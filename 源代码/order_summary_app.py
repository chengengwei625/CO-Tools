# -*- coding: utf-8 -*-
"""订单汇总导出 —— 第三功能（紧随「港口/数量汇总」）。

读取「港口/数量汇总」生成文件中的 `港口明细汇总` 工作表，按 **客款号** 分区、
按 **颜色** 汇总，并把同一颜色在各港口的数量按其 `实际出厂期` 分列填入
（每个交货日期各占一列），生成形如「CO客订单汇总表」的工作簿。

- Style Name（英文款名）从原始 PO CSV 的 `STYLE DESCRIPTION` 按客款号抓取。
- 款式图 / 工厂 / 进度 / 订单类型 / 下单日期 留空给用户手填。

复用 generate_tabs 的：detect_sources / parse_port / latest_purchase_order_csvs
/ _read_po_csv / _split_style_colour / load_saved / update_saved。
"""

import sys
import threading
import traceback
import tkinter as tk
from collections import OrderedDict
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

import generate_tabs as gt

if getattr(sys, 'frozen', False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).parent


# ── 样式 ────────────────────────────────────────────────────────────
_thin = Side(style='thin')
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
HDR_FILL = PatternFill('solid', fgColor='99CCFF')
SUM_FILL = PatternFill('solid', fgColor='FFFFCC')
TITLE_FONT = Font(name='微软雅黑', size=28, bold=True)
TOTAL_FONT = Font(name='微软雅黑', size=20, bold=True)
GRAND_FONT = Font(name='微软雅黑', size=13, bold=True)   # smaller grand-total in G2
HDR_FONT = Font(name='微软雅黑', size=14, bold=True)
DATA_FONT = Font(name='微软雅黑', size=12)
SUM_FONT = Font(name='微软雅黑', size=13, bold=True)
LABEL_FONT = Font(name='微软雅黑', size=11)
CENTER = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center')
DATE_FMT = 'm-d'

FIXED_HDRS = ['款式图', 'Style Name', '客款号', '厂款号', 'Col.code', '颜色', '总件数']
TAIL_HDRS = ['船期', '公司货期', '工厂', '进度', '订单类型', '下单日期']
N_FIXED = len(FIXED_HDRS)        # 7 → A..G
COL_TOTAL = 7                    # G 总件数
HDR_DATE_FMT = 'dd-mmm'          # column-header dates: "23-Jun"
SHIP_DATE_FMT = 'm"月"d"日"'      # 船期/公司货期 cells: "6月23日"


# ── 数据处理 ────────────────────────────────────────────────────────
def _to_date(v):
    """把单元格值解析成 date；解析不了返回 None。"""
    if v is None or v == '':
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = str(v).strip()
    if not s:
        return None
    try:
        return date.fromisoformat(s[:10])
    except ValueError:
        return None


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _int_if_whole(x):
    return int(round(x)) if abs(x - round(x)) < 1e-9 else x


def _code_sort_key(code):
    """客户色号排序键：数字色号按数值升序，非数字排其后按字符串。"""
    s = str(code or '').strip()
    if s.isdigit():
        return (0, int(s), '')
    return (1, 0, s)


def _find_csv_field(row: dict, *candidates: str) -> str:
    """Case-insensitive field lookup across a CSV row dict."""
    lower = {k.lower(): v for k, v in row.items()}
    for c in candidates:
        v = lower.get(c.lower())
        if v is not None:
            return str(v)
    return ''


def load_style_names(po_folder) -> dict:
    """{客款号(str): STYLE DESCRIPTION} —— 从原始 PO CSV 抓取英文款名。

    首先尝试 latest_purchase_order_csvs（按文件名过滤），失败时回退到扫描
    po_folder 下全部 CSV，以应对命名不符合规范的情况。"""
    names: dict[str, str] = {}
    if not po_folder:
        return names
    csvs = []
    try:
        csvs = gt.latest_purchase_order_csvs(po_folder)
    except Exception:                                   # noqa: BLE001
        # Fallback: scan every CSV in the folder
        from pathlib import Path as _P
        csvs = list(_P(po_folder).rglob('*.csv'))
    for p in csvs:
        try:
            rows = gt._read_po_csv(p)
        except Exception:                               # noqa: BLE001
            continue
        for r in rows:
            sc = _find_csv_field(r, 'STYLE COLOUR CODE', 'STYLE COLOR CODE',
                                 'Style Colour Code', 'Style Color Code')
            sd = _find_csv_field(r, 'STYLE DESCRIPTION', 'Style Description')
            style, _ = gt._split_style_colour(sc)
            desc = sd.strip()
            if style and desc and str(style) not in names:
                names[str(style)] = desc
    return names


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


def group_by_style(records: list[dict]) -> "OrderedDict[str, dict]":
    """按客款号分区，区段内按 (客户色号, 中文颜色) 聚合。返回 OrderedDict 客款号 → section。

    section = {'style', 'colors': OrderedDict (色号,中文颜色) → color}
    color   = {'code', 'color', 'fac', 'total', 'by_date'{date:qty}, 'ship', 'company'}

    注1：厂款号按 **每个颜色** 记录 —— 同一客款号可能对应多个厂款号
    （每个色号只属一个厂款号），故不能整区段共用一个厂款号。
    注2：套装(同一客户色号含多个中文颜色)按 **中文颜色** 拆成多行；合并英文串
    （如 "BLACK, COCONUT MILK, PINK KISS"）借助全局 中→英 映射还原为单色英文。
    """
    cn_en_map = _build_cn_en_map(records)
    sections: "OrderedDict[str, dict]" = OrderedDict()
    for rec in records:
        style = str(rec.get('客款号') or '').strip()
        if not style:
            continue
        code = str(rec.get('客户色号') or '').strip()
        cn = str(rec.get('颜色') or '').strip()
        en = str(rec.get('英文颜色') or '').strip()
        # 合并英文串 → 用全局映射还原为该中文色对应的单色英文
        if cn and len(gt._split_colour_list(en)) > 1 and cn in cn_en_map:
            en = cn_en_map[cn]

        sec = sections.get(style)
        if sec is None:
            sec = {'style': style, 'colors': OrderedDict()}
            sections[style] = sec

        ckey = (code, cn)
        col = sec['colors'].get(ckey)
        if col is None:
            # 颜色 = 英文颜色 + 中文颜色；若源数据中文位已是英文（重复）则不拼接
            if cn and cn != en:
                color = f'{en}{cn}'
            else:
                color = en or cn
            col = {'code': code, 'color': color,
                   'fac': str(rec.get('厂款号') or '').strip(),
                   'total': 0.0, 'by_date': {}, 'ship': None, 'company': None}
            sec['colors'][ckey] = col
        if not col['fac']:
            col['fac'] = str(rec.get('厂款号') or '').strip()

        qty = _num(rec.get('总数'))
        col['total'] += qty
        d = _to_date(rec.get('实际出厂期'))
        if d is not None:
            col['by_date'][d] = col['by_date'].get(d, 0.0) + qty
        ship = _to_date(rec.get('实际船期'))
        if ship and (col['ship'] is None or ship < col['ship']):
            col['ship'] = ship
        comp = _to_date(rec.get('工厂出厂期'))
        if comp and (col['company'] is None or comp < col['company']):
            col['company'] = comp
    return sections


# ── 工作簿写出 ──────────────────────────────────────────────────────
def _set(ws, r, c, value, font=DATA_FONT, align=CENTER, border=True,
         fill=None, numfmt=None):
    cell = ws.cell(r, c, value)
    cell.font = font
    cell.alignment = align
    if border:
        cell.border = BORDER
    if fill:
        cell.fill = fill
    if numfmt:
        cell.number_format = numfmt
    return cell


def build_order_summary_workbook(src_path, po_folder, year_quarter):
    """读取港口明细汇总 + PO CSV，构建订单汇总工作簿。返回 (Workbook, sheet_title, n_styles)。"""
    wb_src = load_workbook(src_path, data_only=True)
    _port_ws, qty_ws, suffix, _ = gt.detect_sources(wb_src)
    # 以「数量明细汇总」为准：每行按 (合同号, 客款号, 客户色号) 携带自己的厂款号，
    # 故同一客款号的不同色号可对应不同厂款号（见 6340248 = B54652 + B55088）。
    records = gt.parse_qty(qty_ws)
    wb_src.close()

    style_names = load_style_names(po_folder)
    sections = group_by_style(records)

    # 全局日期列（去重升序）
    all_dates = set()
    for sec in sections.values():
        for col in sec['colors'].values():
            all_dates.update(col['by_date'].keys())
    date_cols = sorted(all_dates)
    n_date = len(date_cols)

    first_date_col = N_FIXED + 1                 # H = 8
    tail_start = first_date_col + n_date         # 船期 起始列
    last_col = tail_start + len(TAIL_HDRS) - 1

    yq = (year_quarter or suffix or '').strip()
    wb = Workbook()
    ws = wb.active
    ws.title = (yq or 'Sheet1')[:31]

    # 第 1 行：更新时间
    ws.cell(1, 1, '更新时间').font = LABEL_FONT
    c = ws.cell(1, 2, date.today())
    c.font = LABEL_FONT
    c.number_format = 'yyyy-m-d'

    # 第 2 行：标题（总件数公式在数据写完后填入）
    ws.cell(2, 1, f'CO {yq}'.strip()).font = TITLE_FONT

    # 第 3 行：表头
    for i, h in enumerate(FIXED_HDRS, start=1):
        _set(ws, 3, i, h, font=HDR_FONT, fill=HDR_FILL)
    for j, d in enumerate(date_cols):
        _set(ws, 3, first_date_col + j, d, font=HDR_FONT, fill=HDR_FILL,
             numfmt=HDR_DATE_FMT)
    for k, h in enumerate(TAIL_HDRS):
        _set(ws, 3, tail_start + k, h, font=HDR_FONT, fill=HDR_FILL)

    # pre-compute the column-letter range for SUM formulas (e.g. "H" to "P")
    first_date_letter = get_column_letter(first_date_col)
    last_date_letter = (get_column_letter(first_date_col + n_date - 1)
                        if n_date > 0 else first_date_letter)

    # 数据区；summary_rows 收集每区段的"合计行"（单色取数据行，多色取小计行）
    r = 4
    summary_rows: list[int] = []
    for sec in sections.values():
        # 按 Col.code(客户色号) 升序；同色号的套装分量保持插入顺序（稳定排序）
        colors = sorted(sec['colors'].values(),
                        key=lambda c: _code_sort_key(c['code']))
        block_first = r
        for col in colors:
            _set(ws, r, 1, None)                                  # 款式图（空）
            _set(ws, r, 2, style_names.get(sec['style'], ''))       # Style Name
            _set(ws, r, 3, sec['style'])                          # 客款号
            _set(ws, r, 4, col['fac'])                            # 厂款号（按色号）
            _set(ws, r, 5, col['code'])                           # Col.code
            _set(ws, r, 6, col['color'])                          # 颜色
            # 总件数 = formula summing all date columns
            if n_date > 0:
                _set(ws, r, COL_TOTAL,
                     f'=SUM({first_date_letter}{r}:{last_date_letter}{r})')
            else:
                _set(ws, r, COL_TOTAL, None)
            for j, d in enumerate(date_cols):
                q = col['by_date'].get(d)
                _set(ws, r, first_date_col + j,
                     _int_if_whole(q) if q else None)
            _set(ws, r, tail_start + 0, col['ship'],
                 numfmt=SHIP_DATE_FMT)                             # 船期
            _set(ws, r, tail_start + 1, col['company'],
                 numfmt=SHIP_DATE_FMT)                             # 公司货期
            for k in range(2, len(TAIL_HDRS)):                     # 工厂/进度/订单类型/下单日期
                _set(ws, r, tail_start + k, None)
            r += 1
        block_last = r - 1

        # 款式图 / Style Name 跨该区段各颜色行合并
        if block_last > block_first:
            ws.merge_cells(start_row=block_first, start_column=1,
                           end_row=block_last, end_column=1)
            ws.merge_cells(start_row=block_first, start_column=2,
                           end_row=block_last, end_column=2)

        # 多颜色区段加「小计」行
        if len(colors) > 1:
            _set(ws, r, 6, '小计', font=SUM_FONT, fill=SUM_FILL)
            _set(ws, r, COL_TOTAL,
                 f'=SUM(G{block_first}:G{block_last})',
                 font=SUM_FONT, fill=SUM_FILL)
            for j in range(n_date):
                cl = get_column_letter(first_date_col + j)
                _set(ws, r, first_date_col + j,
                     f'=SUM({cl}{block_first}:{cl}{block_last})',
                     font=SUM_FONT, fill=SUM_FILL)
            for cc in (1, 2, 3, 4, 5):
                _set(ws, r, cc, None, font=SUM_FONT, fill=SUM_FILL)
            for k in range(len(TAIL_HDRS)):
                _set(ws, r, tail_start + k, None, font=SUM_FONT, fill=SUM_FILL)
            summary_rows.append(r)   # 小计行 = 区段代表
            r += 1
        else:
            summary_rows.append(block_first)  # 单色区段：唯一数据行就是代表

        r += 1   # 区段间空行

    # 第 2 行 G：= 各区段代表行 G 之和（避免双重计入小计）
    if summary_rows:
        refs = ','.join(f'G{row}' for row in summary_rows)
        gc = ws.cell(2, COL_TOTAL, f'=SUM({refs})')
        gc.font = GRAND_FONT
        gc.alignment = CENTER

    _apply_widths(ws, n_date, first_date_col, tail_start)
    return wb, ws.title, len(style_names)


def _apply_widths(ws, n_date, first_date_col, tail_start):
    widths = {1: 26, 2: 34, 3: 15, 4: 15, 5: 14, 6: 28, 7: 13}
    for j in range(n_date):
        widths[first_date_col + j] = 10.5
    for k, w in enumerate((12, 12, 16, 40, 13, 14)):   # 船期/公司货期/工厂/进度/订单类型/下单日期
        widths[tail_start + k] = w
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.row_dimensions[2].height = 36
    ws.row_dimensions[3].height = 24


def default_out_name(year_quarter, title=''):
    """默认输出文件名（不含扩展名）：`{年} CO客订单汇总表 {YYYY.M.D}`。"""
    yq = (year_quarter or title or '').strip()
    year = ''
    for tok in yq.split():
        if tok.isdigit() and len(tok) == 4:
            year = tok
            break
    today = date.today()
    return f'{year} CO客订单汇总表 {today.year}.{today.month}.{today.day}'.strip()


def export_order_summary(src_path, po_folder, year_quarter, out_dir, out_name=''):
    """生成并保存订单汇总，返回 (文件路径, 找到的款式名数量)。

    out_name 为用户自定义文件名（可含或不含 .xlsx）；留空则用默认名。"""
    wb, title, n_styles = build_order_summary_workbook(src_path, po_folder, year_quarter)
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    stem = (out_name or '').strip()
    if stem.lower().endswith('.xlsx'):
        stem = stem[:-5]
    if not stem:
        stem = default_out_name(year_quarter, title)
    fname = _safe_name(stem + '.xlsx')
    path = out / fname
    wb.save(path)
    return str(path), n_styles


def _safe_name(name: str) -> str:
    import re
    name = re.sub(r'[\\/:*?"<>|]', '-', name)
    return re.sub(r'\s+', ' ', name).strip()


# ── UI ──────────────────────────────────────────────────────────────
def build_into(parent):
    cfg = gt.load_saved()
    p = {'padx': 8, 'pady': 4}
    parent.columnconfigure(1, weight=1)

    # Row 0: 港口/数量汇总文件
    tk.Label(parent, text='港口/数量汇总文件：').grid(row=0, column=0, sticky='e', **p)
    src_var = tk.StringVar(value=cfg.get('order_summary_src', ''))
    tk.Entry(parent, textvariable=src_var).grid(row=0, column=1, sticky='ew', **p)

    def _browse_src():
        path = filedialog.askopenfilename(
            filetypes=[('Excel 文件', '*.xlsx'), ('所有文件', '*.*')])
        if path:
            src_var.set(path)
            gt.update_saved(order_summary_src=path)
    tk.Button(parent, text='文件…', command=_browse_src).grid(row=0, column=2, **p)

    # Row 1: PO 合同文件夹（用于 Style Name）
    tk.Label(parent, text='PO 合同文件夹：').grid(row=1, column=0, sticky='e', **p)
    po_var = tk.StringVar(value=cfg.get('po_folder', ''))
    tk.Entry(parent, textvariable=po_var).grid(row=1, column=1, sticky='ew', **p)

    def _browse_po():
        path = filedialog.askdirectory()
        if path:
            po_var.set(path)
            gt.update_saved(po_folder=path)
    tk.Button(parent, text='文件夹…', command=_browse_po).grid(row=1, column=2, **p)

    # Row 2: 年度/季度
    tk.Label(parent, text='年度/季度：').grid(row=2, column=0, sticky='e', **p)
    yq_var = tk.StringVar(value=cfg.get('year_quarter', ''))
    tk.Entry(parent, textvariable=yq_var, width=18).grid(row=2, column=1, sticky='w', **p)

    # Row 3: 输出文件夹
    tk.Label(parent, text='输出文件夹：').grid(row=3, column=0, sticky='e', **p)
    out_var = tk.StringVar(value=cfg.get('order_summary_out', '')
                           or cfg.get('output_waimao', ''))
    tk.Entry(parent, textvariable=out_var).grid(row=3, column=1, sticky='ew', **p)

    def _browse_out():
        path = filedialog.askdirectory()
        if path:
            out_var.set(path)
            gt.update_saved(order_summary_out=path)
    tk.Button(parent, text='文件夹…', command=_browse_out).grid(row=3, column=2, **p)

    # Row 4: 输出文件名（可自定义，留空用默认名）
    tk.Label(parent, text='输出文件名：').grid(row=4, column=0, sticky='e', **p)
    name_var = tk.StringVar(value=default_out_name(yq_var.get().strip()))
    name_frame = tk.Frame(parent)
    name_frame.grid(row=4, column=1, sticky='ew', **p)
    name_frame.columnconfigure(0, weight=1)
    tk.Entry(name_frame, textvariable=name_var).grid(row=0, column=0, sticky='ew')
    tk.Label(name_frame, text='.xlsx', fg='gray').grid(row=0, column=1, padx=(2, 0))

    def _reset_name():
        name_var.set(default_out_name(yq_var.get().strip()))
    tk.Button(parent, text='默认名', command=_reset_name).grid(row=4, column=2, **p)

    tk.Label(parent,
             text='  说明：按客款号分区、按颜色汇总，各交货日期(实际出厂期)各占一列；'
                  'Style Name 取自 PO CSV；款式图/工厂/进度/订单类型/下单日期 留空手填。',
             fg='gray', justify='left').grid(
        row=5, column=0, columnspan=3, sticky='w', padx=8, pady=(2, 6))

    status = tk.StringVar(value='')
    tk.Label(parent, textvariable=status, fg='#005599').grid(
        row=7, column=0, columnspan=3, sticky='w', **p)

    def do_export():
        src = src_var.get().strip()
        out = out_var.get().strip()
        if not src or not Path(src).is_file():
            messagebox.showwarning('提示', '请选择有效的「港口/数量汇总」文件')
            return
        if not out:
            messagebox.showwarning('提示', '请选择输出文件夹')
            return
        # Warn early if PO folder looks invalid (Style Name will be blank)
        po = po_var.get().strip()
        if po and not Path(po).is_dir():
            messagebox.showwarning(
                'PO 文件夹路径无效',
                f'以下 PO 合同文件夹不存在，Style Name 一列将为空白：\n\n{po}\n\n'
                '请点击「文件夹…」重新选择正确的 PO 文件夹后再导出。')
            return
        gt.update_saved(order_summary_src=src, order_summary_out=out,
                        year_quarter=yq_var.get().strip(),
                        po_folder=po)
        export_btn.config(state='disabled')
        status.set('生成中…')
        result = {}

        out_name = name_var.get().strip()

        def worker():
            try:
                result['path'], result['n_styles'] = export_order_summary(
                    src, po, yq_var.get().strip(), out, out_name)
            except Exception:                           # noqa: BLE001
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
            msg = f'已导出：\n{result["path"]}'
            if result.get('n_styles', 1) == 0:
                msg += ('\n\n⚠️  Style Name 列为空白：未从 PO 合同文件夹中找到款式英文名。\n'
                        '请确认 PO 合同文件夹中有 Purchase Order CSV 文件，或重新选择正确路径。')
            status.set('完成')
            messagebox.showinfo('完成', msg)

        parent.after(150, poll)

    btnrow = tk.Frame(parent)
    btnrow.grid(row=6, column=0, columnspan=3, sticky='w', **p)
    export_btn = tk.Button(btnrow, text='  导出订单汇总  ', command=do_export,
                           bg='#2d7d46', fg='white', font=('Arial', 11, 'bold'))
    export_btn.pack(side='left')
    tk.Button(btnrow, text='打开输出文件夹',
              command=lambda: _open_dir(out_var.get())).pack(side='left', padx=10)


def _open_dir(path):
    try:
        if path and Path(path).is_dir():
            import os
            os.startfile(path)                          # noqa: S606
    except Exception:                                   # noqa: BLE001
        pass


def main():
    root = tk.Tk()
    root.title('订单汇总导出（测试版）')
    root.geometry('820x360')
    build_into(root)
    root.mainloop()


if __name__ == '__main__':
    main()
