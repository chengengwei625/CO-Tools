# -*- coding: utf-8 -*-
"""断针记录（金属检针报告）导出 —— 紧随「港口/数量汇总」。

读取「港口/数量汇总」生成文件中的 `港口明细汇总` 工作表，按 **客款号** 生成
一份「断针记录」Word 文档（基于用户提供的模板 docx 填充）：

- STYLE NO. 客款号、PO No. 合同号、DESCRIPTION 产品描述 填入各分节表头；
- 金属检针明细表按 **颜色（客户色号升序）** 分块，每块逐尺码列出「检验总件数」；
- 合格件数 / 退回件数 / 日期 等留空给检验员手填。

填充方式：直接克隆模板里的表格行，保留全部格式（字体、边框、列宽、页眉页脚、
图片），仅替换文字与纵向合并标记 —— 因此输出与模板观感完全一致。
"""

import copy
import sys
import threading
import traceback
import tkinter as tk
from collections import OrderedDict
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import generate_tabs as gt
from order_summary_app import load_style_names

if getattr(sys, 'frozen', False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).parent

XML_SPACE = '{http://www.w3.org/XML/1998/namespace}space'

# 各分节表头中"客款号/合同号/产品描述"所在的 (表序号, 行, 列) —— 见模板结构
STYLE_CELLS = [(0, 0, 1), (3, 1, 1), (9, 0, 1)]   # STYLE NO. 值单元格
PO_CELLS = [(0, 0, 3), (3, 1, 3)]                  # PO No. 值单元格
DESC_CELLS = [(0, 1, 1), (3, 2, 1), (9, 1, 1)]     # DESCRIPTION 值单元格
DETAIL_TABLE = 1                                    # 金属检针明细表


# ── 数据处理 ────────────────────────────────────────────────────────
def _int_if_whole(x):
    return int(round(x)) if abs(x - round(x)) < 1e-9 else x


def _build_cn_en_map(records) -> dict:
    """从单色记录建立 中文颜色→英文颜色 映射，用于把套装的合并英文串还原为单色。"""
    m: dict = {}
    for rec in records:
        cn = str(rec.get('颜色') or '').strip()
        en = str(rec.get('英文颜色') or '').strip()
        if not cn or not en:
            continue
        if len(gt._split_colour_list(en)) != 1 or len(gt._split_colour_list(cn)) != 1:
            continue
        m.setdefault(cn, en)
    return m


def _canonical_sizes(recs) -> list:
    """规范尺码顺序：取最长的 size_headers 作基准，再补入其它未见尺码。"""
    base: list = []
    for r in recs:
        h = r.get('size_headers') or []
        if len(h) > len(base):
            base = list(h)
    order = list(base)
    seen = set(order)
    for r in recs:
        for h in (r.get('size_headers') or []):
            if h not in seen:
                seen.add(h)
                order.append(h)
    return order


def collect_styles(src_path) -> "OrderedDict[str, dict]":
    """返回 OrderedDict 客款号 → {pos[list], sizes[list], blocks[list]}。

    blocks = [(colour_label, [(size, qty_or_None), ...]), ...]  按客户色号升序。
    """
    wb = load_workbook_safe(src_path)
    port_ws, _qty_ws, _suffix, _ = gt.detect_sources(wb)
    records = gt.parse_port(port_ws)
    wb.close()

    cn_en_map = _build_cn_en_map(records)

    by_style: "OrderedDict[str, list]" = OrderedDict()
    for r in records:
        style = str(r.get('客款号') or '').strip()
        if not style:
            continue
        by_style.setdefault(style, []).append(r)

    def _codekey(code):
        try:
            return (0, int(code))
        except (TypeError, ValueError):
            return (1, code)

    out: "OrderedDict[str, dict]" = OrderedDict()
    for style, recs in by_style.items():
        pos: list = []
        for r in recs:
            po = str(r.get('合同号') or '').strip()
            if po and po not in pos:
                pos.append(po)
        sizes = _canonical_sizes(recs)

        # 按 (客户色号, 中文颜色) 聚合 —— 套装(同色号多中文色)拆成多个颜色块；
        # 合并英文串借助全局 中→英 映射还原为单色。
        cols: "OrderedDict[tuple, dict]" = OrderedDict()
        order_idx = 0
        for r in recs:
            code = str(r.get('客户色号') or '').strip()
            cn = str(r.get('颜色') or '').strip()
            en = str(r.get('英文颜色') or '').strip()
            if cn and len(gt._split_colour_list(en)) > 1 and cn in cn_en_map:
                en = cn_en_map[cn]
            ckey = (code, cn)
            c = cols.get(ckey)
            if c is None:
                label = f'{en}{cn}' if (cn and cn != en) else (en or cn)
                c = {'code': code, 'label': label, 'qty': {}, 'order': order_idx}
                order_idx += 1
                cols[ckey] = c
            for h, v in zip(r.get('size_headers') or [], r.get('sizes') or []):
                if v in (None, ''):
                    continue
                try:
                    q = float(v)
                except (TypeError, ValueError):
                    continue
                if q == 0:
                    continue
                c['qty'][h] = c['qty'].get(h, 0.0) + q

        # 按客户色号升序；同色号的套装分量保持插入顺序
        ordered = sorted(cols.values(),
                         key=lambda c: (_codekey(c['code']), c['order']))
        blocks = []
        for c in ordered:
            block = [(s, c['qty'].get(s)) for s in sizes]
            blocks.append((c['label'], block))

        out[style] = {'pos': pos, 'sizes': sizes, 'blocks': blocks}
    return out


def load_workbook_safe(path):
    from openpyxl import load_workbook
    return load_workbook(path, data_only=True)


# ── docx 低层操作 ───────────────────────────────────────────────────
def _set_cell_text(tc, text):
    """把单元格文字设为单一 run，保留段落与字体格式。"""
    paras = tc.findall(qn('w:p'))
    if not paras:
        p = OxmlElement('w:p')
        tc.append(p)
        paras = [p]
    p = paras[0]
    for extra in paras[1:]:
        tc.remove(extra)

    # 取一个 rPr 模板：优先已有 run 的 rPr，否则取段落标记 pPr/rPr
    rpr = None
    existing_r = p.find(qn('w:r'))
    if existing_r is not None:
        rpr = existing_r.find(qn('w:rPr'))
    if rpr is None:
        pPr = p.find(qn('w:pPr'))
        if pPr is not None:
            rpr = pPr.find(qn('w:rPr'))

    for r in p.findall(qn('w:r')):
        p.remove(r)

    if text is None or str(text) == '':
        return
    r = OxmlElement('w:r')
    if rpr is not None:
        r.append(copy.deepcopy(rpr))
    t = OxmlElement('w:t')
    t.set(XML_SPACE, 'preserve')
    t.text = str(text)
    r.append(t)
    p.append(r)


def _set_vmerge(tc, restart: bool):
    """设置/更新单元格的纵向合并标记。"""
    tcPr = tc.find(qn('w:tcPr'))
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        tc.insert(0, tcPr)
    vm = tcPr.find(qn('w:vMerge'))
    if vm is None:
        vm = OxmlElement('w:vMerge')
        tcPr.append(vm)
    if restart:
        vm.set(qn('w:val'), 'restart')
    else:
        if qn('w:val') in vm.attrib:
            del vm.attrib[qn('w:val')]


def _tc_list(tr):
    return tr.findall(qn('w:tc'))


SECTION_HEADING = 'COTTON ON GROUP'


def _force_section_page_breaks(doc):
    """让每个分节标题（COTTON ON GROUP）另起一页（首个除外）。"""
    first = True
    for para in doc.paragraphs:
        if para.text.strip() == SECTION_HEADING:
            if first:
                first = False
                continue
            para.paragraph_format.page_break_before = True


def build_needle_doc(template_path, style, desc, pos, blocks):
    """基于模板生成一份断针记录 Document。"""
    doc = Document(str(template_path))
    po_text = '/'.join(pos)

    def _fill(cells, value):
        for (ti, ri, ci) in cells:
            tbl = doc.tables[ti]
            _set_cell_text(tbl.rows[ri].cells[ci]._tc, value)

    _fill(STYLE_CELLS, style)
    _fill(PO_CELLS, po_text)
    _fill(DESC_CELLS, desc)

    # 每个分节（以「COTTON ON GROUP」开头）都另起一页 —— 因明细表行数可变，
    # 不能再依赖模板原有的固定行数把表头自然顶到下一页。
    _force_section_page_breaks(doc)

    # 重建金属检针明细表的数据行
    tbl = doc.tables[DETAIL_TABLE]
    tbl_el = tbl._tbl
    data_rows = tbl.rows[1:]                    # row0 = 表头
    if not data_rows:
        return doc
    proto = copy.deepcopy(data_rows[0]._tr)     # 含 6 单元格、正确边框/列宽的样板行
    for row in data_rows:
        tbl_el.remove(row._tr)

    for label, block in blocks:
        for j, (size, qty) in enumerate(block):
            tr = copy.deepcopy(proto)
            tcs = _tc_list(tr)
            restart = (j == 0)
            _set_vmerge(tcs[0], restart)        # 颜色（纵向合并）
            _set_vmerge(tcs[1], restart)        # 仪器型号（纵向合并）
            _set_cell_text(tcs[0], label if restart else '')
            _set_cell_text(tcs[1], 'BCT-680D' if restart else '')
            _set_cell_text(tcs[2], size)
            _set_cell_text(tcs[3], _int_if_whole(qty) if qty else '')
            _set_cell_text(tcs[4], '')          # 合格件数 留空
            _set_cell_text(tcs[5], '')          # 退回件数 留空
            tbl_el.append(tr)
    return doc


def _safe_name(name: str) -> str:
    import re
    name = re.sub(r'[\\/:*?"<>|]', '-', name)
    return re.sub(r'\s+', ' ', name).strip()


def export_needle_reports(src_path, po_folder, template_path, out_dir,
                          styles_filter=None) -> list:
    """生成断针记录（每客款号一份），返回写出的文件路径列表。"""
    data = collect_styles(src_path)
    names = load_style_names(po_folder)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)

    wanted = set(styles_filter) if styles_filter else None
    paths = []
    for style, info in data.items():
        if wanted is not None and style not in wanted:
            continue
        desc = names.get(style, '')
        doc = build_needle_doc(template_path, style, desc,
                               info['pos'], info['blocks'])
        fname = _safe_name(f'断针记录 {style}.docx')
        path = out / fname
        doc.save(str(path))
        paths.append(str(path))
    return paths


# ── UI ──────────────────────────────────────────────────────────────
def build_into(parent):
    cfg = gt.load_saved()
    p = {'padx': 8, 'pady': 4}
    parent.columnconfigure(1, weight=1)

    default_tpl = cfg.get('needle_template', '')
    if not default_tpl:
        sample_dir = APP_DIR.parent / '测试数据' / '订单汇总表 sample'
        for guess in (APP_DIR.parent / '模板' / '断针记录模板.docx',
                      sample_dir / '断针记录模板.docx', sample_dir / '断针记录 template.docx'):
            if guess.is_file():
                default_tpl = str(guess)
                break

    # Row 0: 港口/数量汇总文件
    tk.Label(parent, text='港口/数量汇总文件：').grid(row=0, column=0, sticky='e', **p)
    src_var = tk.StringVar(value=cfg.get('needle_src', '')
                           or cfg.get('order_summary_src', ''))
    tk.Entry(parent, textvariable=src_var).grid(row=0, column=1, sticky='ew', **p)

    def _browse_src():
        path = filedialog.askopenfilename(
            filetypes=[('Excel 文件', '*.xlsx'), ('所有文件', '*.*')])
        if path:
            src_var.set(path)
            gt.update_saved(needle_src=path)
            _reload_styles()
    tk.Button(parent, text='文件…', command=_browse_src).grid(row=0, column=2, **p)

    # Row 1: PO 合同文件夹（用于产品描述）
    tk.Label(parent, text='PO 合同文件夹：').grid(row=1, column=0, sticky='e', **p)
    po_var = tk.StringVar(value=cfg.get('po_folder', ''))
    tk.Entry(parent, textvariable=po_var).grid(row=1, column=1, sticky='ew', **p)

    def _browse_po():
        path = filedialog.askdirectory()
        if path:
            po_var.set(path)
            gt.update_saved(po_folder=path)
    tk.Button(parent, text='文件夹…', command=_browse_po).grid(row=1, column=2, **p)

    # Row 2: 模板文件
    tk.Label(parent, text='断针记录模板：').grid(row=2, column=0, sticky='e', **p)
    tpl_var = tk.StringVar(value=default_tpl)
    tk.Entry(parent, textvariable=tpl_var).grid(row=2, column=1, sticky='ew', **p)

    def _browse_tpl():
        path = filedialog.askopenfilename(
            filetypes=[('Word 文档', '*.docx'), ('所有文件', '*.*')])
        if path:
            tpl_var.set(path)
            gt.update_saved(needle_template=path)
    tk.Button(parent, text='文件…', command=_browse_tpl).grid(row=2, column=2, **p)

    # Row 3: 输出文件夹
    tk.Label(parent, text='输出文件夹：').grid(row=3, column=0, sticky='e', **p)
    out_var = tk.StringVar(value=cfg.get('needle_out', '')
                           or cfg.get('output_waimao', ''))
    tk.Entry(parent, textvariable=out_var).grid(row=3, column=1, sticky='ew', **p)

    def _browse_out():
        path = filedialog.askdirectory()
        if path:
            out_var.set(path)
            gt.update_saved(needle_out=path)
    tk.Button(parent, text='文件夹…', command=_browse_out).grid(row=3, column=2, **p)

    # Row 4: 客款号选择
    tk.Label(parent, text='客款号：').grid(row=4, column=0, sticky='e', **p)
    style_var = tk.StringVar(value='【全部】')
    style_cb = ttk.Combobox(parent, textvariable=style_var, state='readonly',
                            values=['【全部】'])
    style_cb.grid(row=4, column=1, sticky='ew', **p)
    tk.Button(parent, text='刷新款号',
              command=lambda: _reload_styles()).grid(row=4, column=2, **p)

    tk.Label(parent,
             text='  说明：按客款号生成断针记录（金属检针报告）；颜色按客户色号升序，'
                  '逐尺码填「检验总件数」；合格/退回件数、日期留空手填；产品描述取自 PO CSV。',
             fg='gray', justify='left').grid(
        row=5, column=0, columnspan=3, sticky='w', padx=8, pady=(2, 6))

    status = tk.StringVar(value='')
    tk.Label(parent, textvariable=status, fg='#005599').grid(
        row=7, column=0, columnspan=3, sticky='w', **p)

    def _reload_styles():
        src = src_var.get().strip()
        if not src or not Path(src).is_file():
            return
        try:
            data = collect_styles(src)
        except Exception:                                   # noqa: BLE001
            status.set('读取款号失败')
            return
        styles = list(data.keys())
        style_cb['values'] = ['【全部】'] + styles
        if style_var.get() not in style_cb['values']:
            style_var.set('【全部】')
        status.set(f'已载入 {len(styles)} 个客款号')

    def do_export():
        src = src_var.get().strip()
        out = out_var.get().strip()
        tpl = tpl_var.get().strip()
        if not src or not Path(src).is_file():
            messagebox.showwarning('提示', '请选择有效的「港口/数量汇总」文件')
            return
        if not tpl or not Path(tpl).is_file():
            messagebox.showwarning('提示', '请选择有效的断针记录模板（.docx）')
            return
        if not out:
            messagebox.showwarning('提示', '请选择输出文件夹')
            return
        po = po_var.get().strip()
        if po and not Path(po).is_dir():
            if not messagebox.askyesno(
                    'PO 文件夹无效',
                    f'PO 合同文件夹不存在，产品描述将为空白：\n{po}\n\n仍要继续吗？'):
                return
        gt.update_saved(needle_src=src, needle_out=out, needle_template=tpl,
                        po_folder=po)
        sel = style_var.get().strip()
        styles_filter = None if sel in ('', '【全部】') else [sel]

        export_btn.config(state='disabled')
        status.set('生成中…')
        result = {}

        def worker():
            try:
                result['paths'] = export_needle_reports(
                    src, po, tpl, out, styles_filter)
            except Exception:                               # noqa: BLE001
                result['err'] = traceback.format_exc()

        t = threading.Thread(target=worker, daemon=True)
        t.start()

        def poll():
            if t.is_alive():
                parent.after(150, poll)
                return
            export_btn.config(state='normal')
            if 'err' in result:
                status.set('生成失败')
                messagebox.showerror('生成失败', result['err'])
                return
            paths = result.get('paths', [])
            status.set(f'完成：生成 {len(paths)} 份')
            if not paths:
                messagebox.showwarning('完成', '未生成任何文件（未匹配到客款号）。')
            else:
                preview = '\n'.join(paths[:10])
                more = '' if len(paths) <= 10 else f'\n… 等共 {len(paths)} 份'
                messagebox.showinfo('完成', f'已导出到：\n{out}\n\n{preview}{more}')

        parent.after(150, poll)

    btnrow = tk.Frame(parent)
    btnrow.grid(row=6, column=0, columnspan=3, sticky='w', **p)
    export_btn = tk.Button(btnrow, text='  生成断针记录  ', command=do_export,
                           bg='#2d7d46', fg='white', font=('Arial', 11, 'bold'))
    export_btn.pack(side='left')
    tk.Button(btnrow, text='打开输出文件夹',
              command=lambda: _open_dir(out_var.get())).pack(side='left', padx=10)

    # 启动时若已有源文件，预载款号
    if src_var.get().strip():
        parent.after(200, _reload_styles)

    return do_export


def _open_dir(path):
    try:
        if path and Path(path).is_dir():
            import os
            os.startfile(path)                              # noqa: S606
    except Exception:                                       # noqa: BLE001
        pass


def main():
    root = tk.Tk()
    root.title('断针记录导出（测试版）')
    root.geometry('860x380')
    build_into(root)
    root.mainloop()


if __name__ == '__main__':
    main()
