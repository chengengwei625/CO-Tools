# -*- coding: utf-8 -*-
"""验货报告填充 —— INLINE(IL 中期) 与 FINAL(FI 最终) 成衣验货报告。

读取「港口/数量汇总」文件的 `数量明细汇总` 工作表，按客款号生成验货报告
（基于用户提供的 .xlsm 模板填充）：

- **INLINE / IL**：每个客款号 2 份（IL1≈25%、IL2≈60% 生产进度）。
- **PASS / FI**：每款一份「未发货」快照，通过「导出FI报告…」勾选导出。

模板含大量图片 + VBA 宏，故**直接改写 zip 内的工作表 XML**（仅替换目标单元格），
其余成员逐字节复制 —— 宏 / 图片 / 公式 / 格式完整保留（openpyxl 仅用于读取）。
"""

import json
import random
import re
import shutil
import sys
import threading
import traceback
import zipfile
import tkinter as tk
from collections import OrderedDict
from datetime import date, datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook

import generate_tabs as gt
import qc_handoff as hf
from order_summary_app import load_style_names

if getattr(sys, 'frozen', False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).parent

# 船期覆盖（用户在程序内手改的最新船期）持久化文件
OVERRIDES_FILE = APP_DIR / 'qc_ship_overrides.json'

# ── 模板「Inspection Report」表的目标单元格 ─────────────────────────
FORM_SHEET_NAME = 'Inspection Report'
CELL_PRODUCT = 'K10'     # 产品描述 Product name
CELL_STYLE = 'K11'       # 款号 Style Number
CELL_ORDERQTY = 'K12'    # 数量 Order Qty（本报告合计）
CELL_LABEL = 'N11'       # IL1/IL2/FI1/FI2…
CELL_SUPPLIER = 'K5'     # 供应商（用于文件名）
CELL_FACTORY = 'K6'      # 厂名（写入交接清单)
CELL_BRAND = 'K7'        # 品牌
CELL_PRODUCTION = 'K8'   # 生产

COL_PO = 'T'             # PO订单号
COL_SHIP = 'V'           # SHIPMENT 发货日期
COL_ORDER = 'X'          # ORDER QTY 数量
COL_CUT = 'Z'            # CUT QTY 已裁剪
COL_PROD = 'AD'          # PRODUCTION 已生产
COL_PACK = 'AJ'          # PACKED QTY 已包装（仅 FI）
DETAIL_FIRST_ROW = 24
DETAIL_LAST_ROW = 33
MAX_DETAIL_ROWS = DETAIL_LAST_ROW - DETAIL_FIRST_ROW + 1   # 10

# ── 其它工作表的「款号 / PO」镜像格 ─────────────────────────────────
# 模板里这些格原本是 ='Inspection Report'!K11 / !T24… 之类的公式，但常被
# QC 同事直接敲字覆盖（实测 Issues / Moisture Control / Fabric-Trims /
# Photos 几乎全被覆盖），重算也救不回来 → 生成时直接写值。
MIRROR_SHEETS = ('Issues', 'Moisture Control', 'Fabric-Trims', 'Packing',
                 'Photos for reports and others')
MIRROR_STYLE_CELL = 'G2'
MIRROR_PRODUCT_CELL = 'AG2'
MIRROR_PO_CELLS = ('L2', 'O2', 'S2', 'V2', 'Y2')        # 这一行只有 5 个 PO 位
MEASURE_SHEET = 'Measurement Chart'
MEASURE_STYLE_CELL = 'B36'
MEASURE_PO_CELLS = tuple('%s36' % c for c in
                         ('F', 'I', 'L', 'O', 'R', 'U', 'X', 'AA', 'AD', 'AG'))

IL_RANGES = (('IL1', 0.20, 0.30), ('IL2', 0.55, 0.65))
FI_PACK_RANGE = (0.85, 0.95)


# ── 通用小工具 ──────────────────────────────────────────────────────
def _to_date(v):
    if v is None or v == '':
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return date.fromisoformat(str(v).strip()[:10])
    except ValueError:
        return None


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def _combine_by_ship(rows):
    """把发货日期相同的行合并为一行：PO 名用「/」连接、数量求和。"""
    agg: "OrderedDict[object, dict]" = OrderedDict()
    for r in rows:
        e = agg.get(r['ship'])
        if e is None:
            e = {'pos': [], 'qty': 0.0}
            agg[r['ship']] = e
        po = str(r.get('po') or '').strip()
        if po and po not in e['pos']:
            e['pos'].append(po)
        e['qty'] += r['qty']
    return [{'po': '/'.join(e['pos']), 'ship': ship, 'qty': e['qty']}
            for ship, e in agg.items()]


def _fit_rows(rows):
    """超过模板 10 行上限时，按发货日期合并行（PO 用「/」连接），否则原样返回。"""
    if len(rows) > MAX_DETAIL_ROWS:
        return _combine_by_ship(rows)
    return rows


def _excel_serial(d) -> int:
    if isinstance(d, datetime):
        d = d.date()
    return (d - date(1899, 12, 30)).days


# ── 船期覆盖（PO 实际船期变动时，用户手填的最新日期）─────────────────
def _ov_key(style, po, ship) -> str:
    """覆盖键 = 客款号||PO||原实际船期ISO（原船期为空时用空串）。"""
    s = ship.isoformat() if isinstance(ship, date) else ''
    return '%s||%s||%s' % (str(style).strip(), str(po or '').strip(), s)


def load_overrides() -> dict:
    try:
        if OVERRIDES_FILE.exists():
            d = json.loads(OVERRIDES_FILE.read_text(encoding='utf-8'))
            if isinstance(d, dict):
                return d
    except Exception:                                       # noqa: BLE001
        pass
    return {}


def save_overrides(d: dict) -> None:
    try:
        OVERRIDES_FILE.write_text(
            json.dumps(d, ensure_ascii=False, indent=2), encoding='utf-8')
    except Exception:                                       # noqa: BLE001
        pass


def _apply_overrides(rows, style, overrides):
    """按原(客款号,PO,原船期)替换 ship 为用户手填的最新日期。"""
    if not overrides:
        return rows
    out = []
    for r in rows:
        nv = overrides.get(_ov_key(style, r['po'], r['ship']))
        nd = _to_date(nv) if nv else None
        out.append({**r, 'ship': nd} if nd else r)
    return out


def _xml_escape(s: str) -> str:
    return (str(s).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;'))


def _safe_name(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', '-', name)
    return re.sub(r'\s+', ' ', name).strip()


def _get_cell_style_idx(xml: str, coord: str) -> 'int | None':
    """Return the s= style index for a cell element, or None if absent."""
    m = _cell_re(coord).search(xml)
    if not m:
        return None
    sm = re.search(r'\bs="(\d+)"', m.group('attrs'))
    return int(sm.group(1)) if sm else None


def _enable_auto_height_rows(xml: str, first_row: int, last_row: int) -> str:
    """Remove fixed-height attrs from <row r="N"> in [first_row, last_row]."""
    def _fix(m):
        tag = m.group(0)
        r_m = re.search(r'\br="(\d+)"', tag)
        if r_m and first_row <= int(r_m.group(1)) <= last_row:
            tag = re.sub(r'\s+ht="[^"]*"', '', tag)
            tag = re.sub(r'\s+customHeight="[^"]*"', '', tag)
        return tag
    return re.sub(r'<row\b[^>]*>', _fix, xml)


def _add_wraptext_to_xf(styles_xml: str, xf_idx: int) -> str:
    """Add wrapText="1" to the cellXfs entry at xf_idx (0-based)."""
    xfs_m = re.search(r'(<cellXfs\b[^>]*>)(.*?)(</cellXfs>)', styles_xml, re.S)
    if not xfs_m:
        return styles_xml
    inner = xfs_m.group(2)
    entries = list(re.finditer(r'<xf\b.*?(?:</xf>|/>)', inner, re.S))
    if xf_idx >= len(entries):
        return styles_xml
    e = entries[xf_idx]
    xf = e.group(0)
    if 'wrapText="1"' in xf:
        return styles_xml
    # Ensure applyAlignment="1"
    if 'applyAlignment' not in xf:
        xf = re.sub(r'(<xf\b)', r'\1 applyAlignment="1"', xf, count=1)
    else:
        xf = xf.replace('applyAlignment="0"', 'applyAlignment="1"')
    # Add wrapText to alignment element or insert one
    if '<alignment' in xf:
        xf = re.sub(r'(<alignment\b)([^>]*?)(/>)', r'\1\2 wrapText="1"\3', xf, count=1)
    elif '</xf>' in xf:
        xf = xf.replace('</xf>', '<alignment wrapText="1"/></xf>', 1)
    else:  # self-closing <xf ... />
        xf = xf[:-2] + '><alignment wrapText="1"/></xf>'
    new_inner = inner[:e.start()] + xf + inner[e.end():]
    return styles_xml[:xfs_m.start(2)] + new_inner + styles_xml[xfs_m.end(2):]


# ── 工作表 XML 单元格写入（保留样式，不碰其它内容）──────────────────
_CELL_RE = {}


def _cell_re(coord):
    pat = _CELL_RE.get(coord)
    if pat is None:
        pat = re.compile(
            r'<c r="%s"(?P<attrs>[^>]*?)(?:/>|>.*?</c>)' % re.escape(coord), re.S)
        _CELL_RE[coord] = pat
    return pat


def _set_cell(xml: str, coord: str, value, kind: str) -> str:
    """替换工作表 XML 中某单元格的值；kind: text|num|date|blank。保留 s= 样式。"""
    m = _cell_re(coord).search(xml)
    if not m:
        raise ValueError('单元格 %s 不在工作表中' % coord)
    sm = re.search(r'\ss="(\d+)"', m.group('attrs'))
    s_attr = ' s="%s"' % sm.group(1) if sm else ''

    if kind == 'blank' or value is None or value == '':
        new = '<c r="%s"%s/>' % (coord, s_attr)
    elif kind == 'text':
        new = ('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s'
               '</t></is></c>' % (coord, s_attr, _xml_escape(value)))
    elif kind == 'date':
        new = '<c r="%s"%s><v>%d</v></c>' % (coord, s_attr, _excel_serial(value))
    else:  # num
        v = value
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        new = '<c r="%s"%s><v>%s</v></c>' % (coord, s_attr, v)
    return xml[:m.start()] + new + xml[m.end():]


def _set_cell_soft(xml: str, coord: str, value, kind: str) -> str:
    """同 _set_cell，但单元格不存在时原样返回（各模板版式略有出入）。"""
    try:
        return _set_cell(xml, coord, value, kind)
    except ValueError:
        return xml


def _fill_mirror(xml, style, product, pos, po_cells, style_cell,
                 product_cell=None) -> str:
    """把款号/产品/PO 直接写进某张镜像表（覆盖公式或残留的手打值）。"""
    xml = _set_cell_soft(xml, style_cell, style,
                         'num' if str(style).isdigit() else 'text')
    if product_cell:
        xml = _set_cell_soft(xml, product_cell, product, 'text')
    for i, c in enumerate(po_cells):
        if i < len(pos):
            xml = _set_cell_soft(xml, c, pos[i], 'text')
        else:
            xml = _set_cell_soft(xml, c, None, 'blank')
    return xml


def _sheet_member(zf: zipfile.ZipFile, name: str) -> str:
    """经 workbook.xml(.rels) 定位某工作表的 XML 成员路径。"""
    wbx = zf.read('xl/workbook.xml').decode('utf-8')
    rid = None
    for mm in re.finditer(r'<sheet\b([^>]+?)/?>', wbx):
        a = mm.group(1)
        nm = re.search(r'name="([^"]+)"', a)
        idm = re.search(r'r:id="(rId\d+)"', a)
        if nm and idm and nm.group(1) == name:
            rid = idm.group(1)
            break
    if rid is None:
        raise ValueError('模板缺少工作表「%s」' % name)
    rels = zf.read('xl/_rels/workbook.xml.rels').decode('utf-8')
    target = None
    for rm in re.finditer(r'<Relationship\b([^>]+?)/?>', rels):
        a = rm.group(1)
        idm = re.search(r'Id="(rId\d+)"', a)
        tgm = re.search(r'Target="([^"]+)"', a)
        if idm and tgm and idm.group(1) == rid:
            target = tgm.group(1)
            break
    if target is None:
        raise ValueError('无法解析工作表关系 %s' % rid)
    if target.startswith('/'):
        return target.lstrip('/')
    return 'xl/' + target


def _form_sheet_member(zf: zipfile.ZipFile) -> str:
    return _sheet_member(zf, FORM_SHEET_NAME)


_WORKBOOK_MEMBER = 'xl/workbook.xml'


def _force_full_calc(xml: str) -> str:
    """让 Excel 打开时全量重算。

    Issues / Moisture Control / Measurement Chart / Pull Test Report 里的
    款号、PO 都是 ='Inspection Report'!K11 之类的公式，只改工作表 XML 不会更新
    它们的**缓存值**——不加这个标志，这些页会一直显示模板里的旧款号/旧 PO。
    """
    m = re.search(r'<calcPr[^>]*/>', xml)
    if not m:
        return xml
    tag = re.sub(r'\s*fullCalcOnLoad="[^"]*"', '', m.group(0))
    return xml[:m.start()] + tag[:-2] + ' fullCalcOnLoad="1"/>' + xml[m.end():]


def fill_report(template_path, out_path, header, rows, label, is_fi) -> int:
    """填充一份验货报告。返回因超过 10 行而被舍弃的行数。"""
    with zipfile.ZipFile(template_path, 'r') as zin:
        member = _form_sheet_member(zin)
        xml = zin.read(member).decode('utf-8')

        # 表头
        xml = _set_cell(xml, CELL_PRODUCT, header['product'], 'text')
        s_kind = 'num' if str(header['style']).isdigit() else 'text'
        xml = _set_cell(xml, CELL_STYLE, header['style'], s_kind)
        xml = _set_cell(xml, CELL_ORDERQTY, header['order_qty'], 'num')
        xml = _set_cell(xml, CELL_LABEL, label, 'text')

        # 清空明细数据行（保留样式）
        for r in range(DETAIL_FIRST_ROW, DETAIL_LAST_ROW + 1):
            for col in (COL_PO, COL_SHIP, COL_ORDER, COL_CUT, COL_PROD, COL_PACK):
                xml = _set_cell(xml, '%s%d' % (col, r), None, 'blank')

        # 写入明细
        n = min(len(rows), MAX_DETAIL_ROWS)
        for i in range(n):
            r = DETAIL_FIRST_ROW + i
            row = rows[i]
            xml = _set_cell(xml, '%s%d' % (COL_PO, r), row['po'], 'text')
            xml = _set_cell(xml, '%s%d' % (COL_SHIP, r), row['ship'], 'date')
            xml = _set_cell(xml, '%s%d' % (COL_ORDER, r), row['qty'], 'num')
            xml = _set_cell(xml, '%s%d' % (COL_CUT, r), row['qty'], 'num')
            xml = _set_cell(xml, '%s%d' % (COL_PROD, r), row['production'], 'num')
            if is_fi:
                xml = _set_cell(xml, '%s%d' % (COL_PACK, r), row['packed'], 'num')

        # Auto row height for detail rows + wrap text on PO column
        po_style_idx = _get_cell_style_idx(xml, '%s%d' % (COL_PO, DETAIL_FIRST_ROW))
        xml = _enable_auto_height_rows(xml, DETAIL_FIRST_ROW, DETAIL_LAST_ROW)
        new_bytes = xml.encode('utf-8')

        # 其它工作表的款号/PO 镜像格：直接写值，不依赖模板公式
        pos = [rows[i]['po'] for i in range(n)]
        mirrors = {}
        for sn in MIRROR_SHEETS + (MEASURE_SHEET,):
            try:
                mm = _sheet_member(zin, sn)
            except ValueError:
                continue
            mx = zin.read(mm).decode('utf-8')
            if sn == MEASURE_SHEET:
                mx = _fill_mirror(mx, header['style'], None, pos,
                                  MEASURE_PO_CELLS, MEASURE_STYLE_CELL)
            else:
                mx = _fill_mirror(mx, header['style'], header['product'], pos,
                                  MIRROR_PO_CELLS, MIRROR_STYLE_CELL,
                                  MIRROR_PRODUCT_CELL)
            mirrors[mm] = mx.encode('utf-8')

        # Patch styles.xml: add wrapText to the PO column's style
        _STYLES_MEMBER = 'xl/styles.xml'
        new_styles = None
        if po_style_idx is not None and _STYLES_MEMBER in zin.namelist():
            try:
                sxml = zin.read(_STYLES_MEMBER).decode('utf-8')
                new_styles = _add_wraptext_to_xf(sxml, po_style_idx).encode('utf-8')
            except Exception:
                pass

        with zipfile.ZipFile(out_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == member:
                    data = new_bytes
                elif item.filename in mirrors:
                    data = mirrors[item.filename]
                elif new_styles and item.filename == _STYLES_MEMBER:
                    data = new_styles
                elif item.filename == _WORKBOOK_MEMBER:
                    data = _force_full_calc(
                        zin.read(item.filename).decode('utf-8')).encode('utf-8')
                else:
                    data = zin.read(item.filename)
                zout.writestr(item, data)
    return len(rows) - n


# ── 数据提取 ────────────────────────────────────────────────────────
def collect_styles(src_path) -> "OrderedDict[str, dict]":
    """返回 OrderedDict 客款号 → {'rows': [{'po','ship','qty'}, ...]}（按 PO+船期聚合）。"""
    wb = load_workbook(src_path, data_only=True)
    _port_ws, qty_ws, _suffix, _ = gt.detect_sources(wb)
    records = gt.parse_qty(qty_ws)
    wb.close()

    by_style: "OrderedDict[str, list]" = OrderedDict()
    for r in records:
        s = str(r.get('客款号') or '').strip()
        if not s:
            continue
        by_style.setdefault(s, []).append(r)

    out: "OrderedDict[str, dict]" = OrderedDict()
    for style, recs in by_style.items():
        agg: "OrderedDict[tuple, float]" = OrderedDict()
        for r in recs:
            po = str(r.get('合同号') or '').strip()
            shipd = _to_date(r.get('实际船期'))
            key = (po, shipd)
            agg[key] = agg.get(key, 0.0) + _num(r.get('总数'))
        rows = [{'po': po, 'ship': shipd, 'qty': q}
                for (po, shipd), q in agg.items() if q]
        out[style] = {'rows': rows}
    return out


def _template_meta(path) -> dict:
    """从模板读取 供应商/品牌/生产（用于文件名）。"""
    meta = {'supplier': '', 'factory': '', 'brand': '', 'production': ''}
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb[FORM_SHEET_NAME] if FORM_SHEET_NAME in wb.sheetnames else wb.active
        meta['supplier'] = str(ws[CELL_SUPPLIER].value or '').strip()
        meta['factory'] = str(ws[CELL_FACTORY].value or '').strip()
        meta['brand'] = str(ws[CELL_BRAND].value or '').strip()
        meta['production'] = str(ws[CELL_PRODUCTION].value or '').strip()
        wb.close()
    except Exception:                                       # noqa: BLE001
        pass
    return meta


def _date_token(d) -> str:
    """文件名日期，如 '7 APR 26'。"""
    if d is None:
        d = date.today()
    return '%d %s %s' % (d.day, d.strftime('%b').upper(), d.strftime('%y'))


def _filename(meta, prefix, style, label, d, inspector) -> str:
    parts = [prefix]
    if meta.get('supplier'):
        parts.append(meta['supplier'])
    parts.append('Style %s' % style)
    bp = ' '.join(x for x in (meta.get('brand'), meta.get('production')) if x)
    if bp:
        parts.append(bp)
    tail = '%s-%s' % (label, _date_token(d))
    if inspector:
        tail += '-%s' % inspector
    parts.append(tail)
    return _safe_name(' - '.join(parts)) + '.xlsm'


def _detect_batch(src_path, cutoff):
    """批次 = 季度（Q1..Q4）。优先取自「港口/数量汇总」源文件——先看工作表名后缀
    （如 '2026 Q2' / '2026Q2'），再看文件名（如 'CO 2026 Q2 …'）；都没有则按 cutoff
    月份推算。返回 (年份字符串, 'Q#')，作为 输出根/<年份>/<季度> 的两级批次目录。"""
    suffix = ''
    try:
        wb = load_workbook(src_path, read_only=True)
        try:
            _p, _q, suffix, _ = gt.detect_sources(wb)
        finally:
            wb.close()
    except Exception:                                       # noqa: BLE001
        suffix = ''
    hay = '%s %s' % (suffix, Path(src_path).name)
    mq = re.search(r'[Qq]\s*([1-4])', hay)
    quarter = 'Q%s' % mq.group(1) if mq else 'Q%d' % ((cutoff.month - 1) // 3 + 1)
    my = re.search(r'(20\d{2})', hay)
    year = my.group(1) if my else str(cutoff.year)
    return year, quarter


def export_qc_reports(src, po_folder, inline_tpl, pass_tpl, out_dir,
                      inspector='', styles_filter=None, cutoff=None,
                      ship_overrides=None, bom_src=None, needle_template=None,
                      mode='both') -> dict:
    """生成 IL/FI 验货报告到「批次文件夹」并写交接清单 manifest.json。

    输出结构: out_dir/<批次>/  含 .xlsm 报告 + BOM + manifest.json,供 QC 工具一键载入。
    返回 {'files': [...], 'overflow': [...], 'batch_dir': str, 'entries': [...]}。

    cutoff        —— 验货(截止)日期；FI 只含未发货 PO（实际船期为空或 ≥ cutoff）。
    ship_overrides—— {覆盖键: 新船期ISO}，套用用户手填的最新船期。
    bom_src       —— BOM物料分类汇总.xlsx 路径(若存在则拷入批次并记入 manifest)。
    mode          —— 'il'|'fi'|'both'，控制本次生成 IL/FI 中的哪一种（默认两者都生成）。
    """
    if cutoff is None:
        cutoff = date.today()
    ship_overrides = ship_overrides or {}
    data = collect_styles(src)
    names = load_style_names(po_folder)
    il_meta = _template_meta(inline_tpl)
    fi_meta = _template_meta(pass_tpl)

    # 批次 = 季度(Q1..Q4)，取自源「港口/数量汇总」；按 年份/季度 归档
    # (out_dir 即共享盘「CO查货资料」根目录)。同一季度重复导出归入同一批次文件夹。
    batch_year, batch_id = _detect_batch(src, cutoff)
    out = Path(out_dir) / batch_year / batch_id
    out.mkdir(parents=True, exist_ok=True)

    wanted = set(styles_filter) if styles_filter else None
    result = {'files': [], 'overflow': [], 'batch_dir': str(out), 'entries': []}

    def _entry(path, meta, style, mode, label, total, produced):
        e = {'file': path.name, 'style': style, 'mode': mode, 'label': label,
             'supplier': meta.get('supplier', ''), 'factory': meta.get('factory', ''),
             'total_order_qty': int(total), 'produced_qty': int(produced)}
        result['entries'].append(e)

    for style, info in data.items():
        if wanted is not None and style not in wanted:
            continue
        product = names.get(style, '')
        all_rows = _apply_overrides(info['rows'], style, ship_overrides)
        if not all_rows:
            continue
        il_date = min((r['ship'] for r in all_rows if r['ship']), default=None)

        # ── INLINE：IL1 / IL2（取该款全部 PO；中期报告早于发货，不按未发货过滤）──
        if mode in ('il', 'both'):
            il_rows = _fit_rows(all_rows)
            for label, lo, hi in IL_RANGES:
                rows = [{**r, 'production': int(round(r['qty'] * random.uniform(lo, hi)))}
                        for r in il_rows]
                total = int(round(sum(r['qty'] for r in il_rows)))
                produced = int(sum(r['production'] for r in rows))
                header = {'product': product, 'style': style, 'order_qty': total}
                path = out / _filename(il_meta, 'INLINE', style, label, il_date, inspector)
                ov = fill_report(inline_tpl, path, header, rows, label, is_fi=False)
                result['files'].append(str(path))
                _entry(path, il_meta, style, 'inline', label, total, produced)
                if ov > 0:
                    result['overflow'].append((path.name, ov))

        # ── FINAL：每款一份，列出所有未发货 PO（实际船期为空或 ≥ 截止日）──
        if mode in ('fi', 'both'):
            unshipped = [r for r in all_rows if r['ship'] is None or r['ship'] >= cutoff]
            if unshipped:
                frows = _fit_rows(unshipped)
                rows = [{**r, 'production': int(round(r['qty'])),
                         'packed': int(round(r['qty'] * random.uniform(*FI_PACK_RANGE)))}
                        for r in frows]
                ftotal = int(round(sum(r['qty'] for r in frows)))
                produced = int(sum(r['production'] for r in rows))
                header = {'product': product, 'style': style, 'order_qty': ftotal}
                # 文件名按截止日期打戳：不同截止日 → 不同文件，旧文件保留
                path = out / _filename(fi_meta, 'PASS', style, 'FI', cutoff, inspector)
                ov = fill_report(pass_tpl, path, header, rows, 'FI', is_fi=True)
                result['files'].append(str(path))
                _entry(path, fi_meta, style, 'final', 'FI', ftotal, produced)
                if ov > 0:
                    result['overflow'].append((path.name, ov))

    # ── 断针记录(每客款号一份 .docx,随批次,供 QC 打印手填后拍照上传)────
    if needle_template and Path(needle_template).is_file():
        try:
            import needle_report_app as nr
            batch_styles = sorted({e['style'] for e in result['entries']})
            nr.export_needle_reports(src, po_folder, needle_template, str(out),
                                     styles_filter=batch_styles)
            # 同款的 IL1/IL2/FI 指向同一份断针记录
            for e in result['entries']:
                cand = out / nr._safe_name('断针记录 %s.docx' % e['style'])
                e['needle_file'] = cand.name if cand.is_file() else None
        except Exception:                                   # noqa: BLE001
            traceback.print_exc()

    # ── 交接清单 + BOM ──────────────────────────────────────────────
    bom_name = None
    if bom_src and Path(bom_src).is_file():
        bom_name = Path(bom_src).name
        try:
            shutil.copy2(bom_src, out / bom_name)
        except OSError:
            bom_name = None
    old = hf.read_manifest(str(out)) or {}
    new_files = {e['file'] for e in result['entries']}
    old_reports = [r for r in old.get('reports', []) if r.get('file') not in new_files]
    merged = old_reports + result['entries']
    bom_final = bom_name or old.get('bom_file')
    hf.write_manifest(str(out), hf.build_manifest(batch_id, merged, bom_final))

    return result


# ── UI ──────────────────────────────────────────────────────────────
def build_into(parent):
    cfg = gt.load_saved()
    p = {'padx': 8, 'pady': 4}
    parent.columnconfigure(1, weight=1)

    def _row(idx, label, key, browse, default=''):
        tk.Label(parent, text=label).grid(row=idx, column=0, sticky='e', **p)
        var = tk.StringVar(value=cfg.get(key, '') or default)
        tk.Entry(parent, textvariable=var).grid(row=idx, column=1, sticky='ew', **p)
        if browse:
            tk.Button(parent, text=browse[0],
                      command=lambda: browse[1](var)).grid(row=idx, column=2, **p)
        return var

    def _pick_file(var, cfgkey):
        path = filedialog.askopenfilename(
            filetypes=[('Excel 文件', '*.xlsx;*.xlsm'), ('所有文件', '*.*')])
        if path:
            var.set(path)
            gt.update_saved(**{cfgkey: path})

    def _pick_dir(var, cfgkey):
        path = filedialog.askdirectory()
        if path:
            var.set(path)
            gt.update_saved(**{cfgkey: path})

    src_var = _row(0, '港口/数量汇总文件：', 'qc_src',
                   ('文件…', lambda v: (_pick_file(v, 'qc_src'), _reload_styles())))
    po_var = _row(1, 'PO 合同文件夹：', 'po_folder',
                  ('文件夹…', lambda v: _pick_dir(v, 'po_folder')))
    il_var = _row(2, 'INLINE 模板(.xlsm)：', 'qc_inline_template',
                  ('文件…', lambda v: _pick_file(v, 'qc_inline_template')))
    fi_var = _row(3, 'PASS 模板(.xlsm)：', 'qc_pass_template',
                  ('文件…', lambda v: _pick_file(v, 'qc_pass_template')))
    out_var = _row(4, '输出文件夹：', 'qc_out',
                   ('文件夹…', lambda v: _pick_dir(v, 'qc_out')),
                   default=cfg.get('output_waimao', ''))
    insp_var = _row(5, '验货员(可选)：', 'qc_inspector', None)

    # 验货(截止)日期 + 船期编辑入口；FI 只含此日期当天或之后未发货的 PO
    cutoff_var = _row(6, '验货日期(截止)：', 'qc_cutoff',
                      ('编辑未发货PO船期…', lambda v: _open_edit_dialog()),
                      default=date.today().isoformat())

    # 用户手填的最新船期覆盖（持久化到 qc_ship_overrides.json）
    overrides = load_overrides()

    # 客款号选择
    tk.Label(parent, text='客款号：').grid(row=7, column=0, sticky='e', **p)
    style_var = tk.StringVar(value='【全部】')
    style_cb = ttk.Combobox(parent, textvariable=style_var, state='readonly',
                            values=['【全部】'])
    style_cb.grid(row=7, column=1, sticky='ew', **p)
    tk.Button(parent, text='刷新款号',
              command=lambda: _reload_styles()).grid(row=7, column=2, **p)

    tk.Label(parent,
             text='  说明：IL 中期报告每款 2 份(IL1≈25%、IL2≈60%，含全部 PO)；'
                  '「生成验货报告(全部IL)」导出全部 IL；FI 请点「导出FI报告…」'
                  '按截止日期勾选未发货的款导出；两者写入同一季度批次并合并交接清单。'
                  '船期变动可在「编辑未发货PO船期…」里手改；文件名按验货日期打戳，每次另存。'
                  '明细表上限 10 行；产品名取自 PO CSV；只填指定字段，其余保留模板原样。',
             fg='gray', justify='left', wraplength=820).grid(
        row=8, column=0, columnspan=3, sticky='w', padx=8, pady=(2, 6))

    status = tk.StringVar(value='')
    tk.Label(parent, textvariable=status, fg='#005599').grid(
        row=10, column=0, columnspan=3, sticky='w', **p)

    def _parse_cutoff():
        d = _to_date(cutoff_var.get().strip())
        return d or date.today()

    def _open_edit_dialog():
        src = src_var.get().strip()
        if not src or not Path(src).is_file():
            messagebox.showwarning('提示', '请先选择有效的「港口/数量汇总」文件')
            return
        try:
            data = collect_styles(src)
        except Exception:                                   # noqa: BLE001
            messagebox.showerror('读取失败', traceback.format_exc())
            return
        cutoff = _parse_cutoff()
        sel = style_var.get().strip()
        wanted = None if sel in ('', '【全部】') else sel
        work = dict(overrides)                              # 编辑副本，保存才生效

        win = tk.Toplevel(parent)
        win.title('编辑未发货 PO 船期')
        win.geometry('780x480')
        win.transient(parent)

        tk.Label(win, fg='gray', justify='left', wraplength=760,
                 text='截止(验货)日期：%s　·　实际船期为空或 ≥ 截止日 → 未发货(计入 FI)；'
                      '早于截止日 → 已发货(不计入)。\n双击「当前船期」修改，'
                      '回车确认（格式 YYYY-MM-DD；留空=恢复原船期）。' % cutoff.isoformat()
                 ).pack(side='top', fill='x', padx=8, pady=(8, 4))

        body = tk.Frame(win)
        body.pack(side='top', fill='both', expand=True)
        cols = ('style', 'po', 'ship', 'qty', 'state')
        tree = ttk.Treeview(body, columns=cols, show='headings', height=15)
        for c, t, w in (('style', '客款号', 120), ('po', 'PO(合同号)', 160),
                        ('ship', '当前船期', 120), ('qty', '数量', 90),
                        ('state', '状态', 90)):
            tree.heading(c, text=t)
            tree.column(c, width=w, anchor='center')
        vsb = ttk.Scrollbar(body, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side='left', fill='both', expand=True, padx=(8, 0), pady=4)
        vsb.pack(side='left', fill='y', pady=4)

        meta = {}                                           # iid -> (style, po, orig_ship)

        def _eff_ship(style, po, orig):
            nv = work.get(_ov_key(style, po, orig))
            nd = _to_date(nv) if nv else None
            return nd if nd else orig

        def _state(ship):
            return '未发货' if (ship is None or ship >= cutoff) else '已发货'

        for style, sinfo in data.items():
            if wanted is not None and style != wanted:
                continue
            for r in sinfo['rows']:
                orig = r['ship']
                eff = _eff_ship(style, r['po'], orig)
                iid = tree.insert('', 'end', values=(
                    style, r['po'], eff.isoformat() if eff else '',
                    int(round(r['qty'])), _state(eff)))
                meta[iid] = (style, r['po'], orig)

        editor = {'entry': None}

        def _cancel_edit(_=None):
            if editor['entry'] is not None:
                editor['entry'].destroy()
                editor['entry'] = None

        def _begin_edit(event):
            _cancel_edit()
            iid = tree.identify_row(event.y)
            if not iid or tree.identify_column(event.x) != '#3':  # 仅「当前船期」列
                return
            bbox = tree.bbox(iid, '#3')
            if not bbox:
                return
            x, y, w, h = bbox
            e = tk.Entry(tree)
            e.insert(0, tree.set(iid, 'ship'))
            e.place(x=x, y=y, width=w, height=h)
            e.focus_set()
            e.select_range(0, 'end')
            editor['entry'] = e

            def _commit(_=None):
                val = e.get().strip()
                style, po, orig = meta[iid]
                key = _ov_key(style, po, orig)
                if val == '':
                    work.pop(key, None)                      # 留空 = 恢复原船期
                    nd = orig
                else:
                    nd = _to_date(val)
                    if nd is None:
                        messagebox.showwarning('日期格式',
                                               '请输入 YYYY-MM-DD（留空=恢复原船期）',
                                               parent=win)
                        return
                    if nd == orig:
                        work.pop(key, None)
                    else:
                        work[key] = nd.isoformat()
                _cancel_edit()
                tree.set(iid, 'ship', nd.isoformat() if nd else '')
                tree.set(iid, 'state', _state(nd))

            e.bind('<Return>', _commit)
            e.bind('<Escape>', _cancel_edit)
            e.bind('<FocusOut>', _cancel_edit)

        tree.bind('<Double-1>', _begin_edit)

        def _save_close():
            _cancel_edit()
            overrides.clear()
            overrides.update(work)
            save_overrides(overrides)
            status.set('已保存船期覆盖 %d 条' % len(overrides))
            win.destroy()

        btns = tk.Frame(win)
        btns.pack(side='bottom', fill='x', pady=6)
        tk.Button(btns, text='保存并关闭', command=_save_close,
                  bg='#2d7d46', fg='white').pack(side='right', padx=8)
        tk.Button(btns, text='取消', command=win.destroy).pack(side='right')

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
        status.set('已载入 %d 个客款号' % len(styles))

    def _start_export(mode, styles_filter):
        src = src_var.get().strip()
        il = il_var.get().strip()
        fi = fi_var.get().strip()
        out = out_var.get().strip()
        if not src or not Path(src).is_file():
            messagebox.showwarning('提示', '请选择有效的「港口/数量汇总」文件')
            return
        if mode in ('il', 'both') and (not il or not Path(il).is_file()):
            messagebox.showwarning('提示', '请选择有效的 INLINE 模板(.xlsm)')
            return
        if mode in ('fi', 'both') and (not fi or not Path(fi).is_file()):
            messagebox.showwarning('提示', '请选择有效的 PASS 模板(.xlsm)')
            return
        if not out:
            messagebox.showwarning('提示', '请选择输出文件夹')
            return
        po = po_var.get().strip()
        if po and not Path(po).is_dir():
            if not messagebox.askyesno(
                    'PO 文件夹无效',
                    'PO 合同文件夹不存在，产品描述将为空白：\n%s\n\n仍要继续吗？' % po):
                return
        cutoff = _parse_cutoff()
        gt.update_saved(qc_src=src, qc_inline_template=il, qc_pass_template=fi,
                        qc_out=out, qc_inspector=insp_var.get().strip(), po_folder=po,
                        qc_cutoff=cutoff.isoformat())

        # BOM(若已在「BOM分类」生成)随批次一并拷入,供 QC 端自动关联水分报告
        bom_dir = gt.load_saved().get('bom_out', '')
        bom_src = str(Path(bom_dir) / hf.BOM_FILENAME) if bom_dir else ''
        # 断针记录模板(在「断针记录」页配置):随批次生成每款一份 .docx 供打印手填
        needle_tpl = gt.load_saved().get('needle_template', '')

        export_btn.config(state='disabled')
        status.set('生成中…')
        result = {}

        def worker():
            try:
                result['data'] = export_qc_reports(
                    src, po, il, fi, out, insp_var.get().strip(), styles_filter,
                    cutoff=cutoff, ship_overrides=overrides, bom_src=bom_src,
                    needle_template=needle_tpl, mode=mode)
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
            data = result['data']
            files = data['files']
            status.set('完成：生成 %d 份(批次已写交接清单)' % len(files))
            if not files:
                messagebox.showwarning('完成', '未生成任何文件（未匹配到客款号或无数据）。')
                return
            batch_dir = data.get('batch_dir', out)
            bom_note = '' if any(Path(batch_dir).glob(hf.BOM_FILENAME)) else \
                '\n\n⚠️  本批次未含 BOM(请先在「BOM分类」生成),QC 端水分报告需手动选 BOM。'
            msg = ('已导出 %d 份到批次文件夹：\n%s\n\n已写入 manifest.json,'
                   'QC 端可「从交接载入」一键打开。' % (len(files), batch_dir)) + bom_note
            if data['overflow']:
                lines = '\n'.join('  %s：超出 %d 行' % (n, c) for n, c in data['overflow'])
                msg += ('\n\n⚠️  以下报告 (PO,船期) 行数超过模板 10 行上限，多余行未填入：\n'
                        + lines)
            messagebox.showinfo('完成', msg)

        parent.after(150, poll)

    def do_export():
        sel = style_var.get().strip()
        styles_filter = None if sel in ('', '【全部】') else [sel]
        _start_export('il', styles_filter)

    def _open_fi_dialog():
        src = src_var.get().strip()
        if not src or not Path(src).is_file():
            messagebox.showwarning('提示', '请先选择有效的「港口/数量汇总」文件')
            return
        cutoff = _parse_cutoff()
        try:
            data = collect_styles(src)
        except Exception:                                   # noqa: BLE001
            messagebox.showerror('读取失败', traceback.format_exc())
            return

        rows_by_style = {}
        for style, info in data.items():
            rows = _apply_overrides(info['rows'], style, overrides)
            unshipped = [r for r in rows if r['ship'] is None or r['ship'] >= cutoff]
            if unshipped:
                rows_by_style[style] = unshipped

        if not rows_by_style:
            messagebox.showinfo('提示', '没有未发货的款（截止 %s）' % cutoff.isoformat())
            return

        win = tk.Toplevel(parent)
        win.title('导出 FI 报告')
        win.geometry('700x420')
        win.transient(parent)

        tk.Label(win, fg='gray', justify='left', wraplength=680,
                 text='截止日期 %s · 单击「☑」列勾选要导出的 FI 款；表头☑全选/全不选'
                      % cutoff.isoformat()
                 ).pack(side='top', fill='x', padx=8, pady=(8, 4))

        body = tk.Frame(win)
        body.pack(side='top', fill='both', expand=True)
        cols = ('chk', 'style', 'pos', 'qty', 'first', 'last')
        tree = ttk.Treeview(body, columns=cols, show='headings', height=14)
        widths = {'chk': 40, 'style': 120, 'pos': 100, 'qty': 110,
                  'first': 110, 'last': 110}
        titles = {'chk': '☑', 'style': '客款号', 'pos': '未发货PO数',
                  'qty': '未发货数量', 'first': '最早船期', 'last': '最晚船期'}
        for c in cols:
            tree.heading(c, text=titles[c])
            tree.column(c, width=widths[c], anchor='center')
        vsb = ttk.Scrollbar(body, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side='left', fill='both', expand=True, padx=(8, 0), pady=4)
        vsb.pack(side='left', fill='y', pady=4)

        for style, rows in rows_by_style.items():
            ships = [r['ship'] for r in rows if r['ship'] is not None]
            first = min(ships).isoformat() if ships else ''
            last = max(ships).isoformat() if ships else ''
            qty = int(round(sum(r['qty'] for r in rows)))
            tree.insert('', 'end', values=('☐', style, len(rows), qty, first, last))

        def _toggle_all_chk():
            items = tree.get_children()
            if not items:
                return
            all_checked = all(tree.set(i, 'chk') == '☑' for i in items)
            new = '☐' if all_checked else '☑'
            for i in items:
                tree.set(i, 'chk', new)

        tree.heading('chk', text='☑', command=_toggle_all_chk)

        def _on_click(event):
            if tree.identify_column(event.x) != '#1':
                return
            row_id = tree.identify_row(event.y)
            if not row_id:
                return
            cur = tree.set(row_id, 'chk')
            tree.set(row_id, 'chk', '☐' if cur == '☑' else '☑')

        tree.bind('<Button-1>', _on_click)

        def _do_export_selected():
            checked = [tree.set(i, 'style') for i in tree.get_children()
                       if tree.set(i, 'chk') == '☑']
            if not checked:
                messagebox.showwarning('提示', '请先勾选至少一款', parent=win)
                return
            win.destroy()
            _start_export('fi', checked)

        btns = tk.Frame(win)
        btns.pack(side='bottom', fill='x', pady=6)
        tk.Button(btns, text='导出所选 FI', command=_do_export_selected,
                  bg='#2d7d46', fg='white').pack(side='right', padx=8)
        tk.Button(btns, text='取消', command=win.destroy).pack(side='right')

    btnrow = tk.Frame(parent)
    btnrow.grid(row=9, column=0, columnspan=3, sticky='w', **p)
    export_btn = tk.Button(btnrow, text='  生成验货报告(全部IL)  ', command=do_export,
                           bg='#2d7d46', fg='white', font=('Arial', 11, 'bold'))
    export_btn.pack(side='left')
    tk.Button(btnrow, text='导出FI报告…',
              command=_open_fi_dialog).pack(side='left', padx=10)
    tk.Button(btnrow, text='打开输出文件夹',
              command=lambda: _open_dir(out_var.get())).pack(side='left', padx=10)

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
    root.title('验货报告生成（测试版）')
    root.geometry('880x420')
    build_into(root)
    root.mainloop()


if __name__ == '__main__':
    main()
