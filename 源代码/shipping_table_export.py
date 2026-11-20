"""
出货表导出工具 —— 第四功能（独立测试版）

输入：第三功能生成的「装箱单（PL）」xlsx 文件夹。
输出：按「目的地分组 × 出货日期」各生成一个「出货表」xlsx（版式仿 测试数据/sample 1）。

要点：
- 生成的 PL 里 CBM/净重/毛重/总件数 等都是 *公式*（openpyxl data_only 读不到值），
  所以本工具只读取原始单元格（箱号区间 B/D、各尺码件数 F:X、箱规 AA/AB/AC、
  纸箱重量 AG、单件净重 AJ10），自行重算 件数 / CBM / 净重 / 毛重。
- 出货日期来自 PL Z3（写入 packing_list_app 时放在 Z3:AI3），格式如 "6.5"。
  同一分组内，相同出货日期的所有 PO 合入一个文件；到深圳日期自动 +1 天。
- 目的地分组在运行时决定（见 GroupSpec / export_shipping_tables）。
- 暂未由数据提供的字段：PO#订单号 / 成分 留空，品牌取常量，编号自动顺延。

可单独运行：
    python shipping_table_export.py <PL文件夹> <输出文件夹> [--groups groups.json]
不带 --groups 时，每个目的地各出一组（按日期拆文件）。
"""

from __future__ import annotations

import json
import re
import sys
from collections import OrderedDict
from dataclasses import dataclass, field
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── 默认常量 ─────────────────────────────────────────────────────────
DEFAULT_BRAND = 'COTTON:ON BODY'
DEFAULT_FACTORY_ADDR = '1. 加工厂地址: 普宁市流沙中河东侧华溪村南华路华都北区2号楼光明楼'
DEFAULT_CONTACT = '联系人：朱生'
DEFAULT_PHONE = '联系电话：18122678828'
DEFAULT_REPORT_NOTE = '***以上合同用联集的文件报关， 谢谢！***'
DEFAULT_QUESTION_NOTE = '如有任何疑问请联系集杰陈小姐(18029962044) '
DEFAULT_HANGER_NOTE = '一件文胸配一个衣架'
DEFAULT_MAKER_TO = 'TO:联集'

WEEKDAYS = ['周一', '周二', '周三', '周四', '周五', '周六', '周日']

# 点箱数明细 顶部固定说明（取自参考文件 26.5.31点箱数明细.xlsx）
DIANKA_NOTE1 = ('注：装车前须核对此份资料的数量是否与出货表符合，'
                '如有异常请告知跟单核对无误后再进行装车！')
DIANKA_NOTE2 = ('散货装车须知：\n'
                '1.请按照目的港、订单号、客款号、及颜色按顺序装车，'
                '对于同一目的港产品未集中堆放或产品与其他产品混放，仓库会收取分货费。\n'
                '2、请按照以下装车次序进行装车，先卸的款号写在第一行，依次类推，'
                '对于同一产品未集中堆放或产品与其他产品混放，仓库会收取分货费。')

# ── 出货表列号（1 基，已去掉 Maersk SO#/包装组/QC） ──────────────────
#   A     B      C        D      E       F      G     H      I
# tag  合同号  客款号   厂款号  成分  包装方式  数量  总箱数  箱数
#   J   K   L   M   N      O     P     Q     R      S     T
#   长   X   宽   X   高   CBM  净重   毛重  总净重  总毛重  品牌
COL = {
    'tag': 1, '合同号': 2, '客款号': 3, '厂款号': 4,
    '成分': 5, '包装方式': 6, '数量': 7, '总箱数': 8, '箱数': 9,
    '长': 10, 'X1': 11, '宽': 12, 'X2': 13, '高': 14, 'CBM': 15,
    '净重': 16, '毛重': 17, '总净重': 18, '总毛重': 19, '品牌': 20,
}

# PL 表内单元格位置（1 基行列）
PL = {
    '合同号': (3, 4),    # D3
    '客款号': (5, 4),    # D5
    '目的地英': (2, 7),  # G2
    '目的地中': (3, 7),  # G3
    'NW': (7, 36),       # AJ7
    '出货日期': (3, 26), # Z3
}
PL_BOX_FIRST_ROW = 14
PL_C_COLOR = 1           # A 颜色（中/英，箱号区）
PL_C_BOXNO_START = 2     # B
PL_C_BOXNO_END = 4       # D
PL_C_SIZE_FIRST = 6      # F
PL_C_SIZE_LAST = 24      # X
PL_C_LEN = 27            # AA 长
PL_C_WID = 28            # AB 宽
PL_C_HGT = 29            # AC 高
PL_C_CARTON = 33         # AG 纸箱重量

# PL 底部「色号对照表」：表头行 col1=主布颜色 col2=客色号 col5=色号
PL_CMAP_HDR = '主布颜色'
PL_CMAP_C_CN = 1         # A 主布颜色（中文）
PL_CMAP_C_EN = 2         # B 客色号（英文）
PL_CMAP_C_CODE = 5       # E 色号


# ── 数据模型 ────────────────────────────────────────────────────────
@dataclass
class Tier:
    """同一箱规（长×宽×高）的一档。"""
    length: int
    width: int
    height: int
    boxes: int = 0
    pieces: int = 0
    net: float = 0.0
    gross: float = 0.0
    cbm: float = 0.0


@dataclass
class ColorBoxes:
    """点箱数明细用：同一 PO 下同一色号的各箱规箱数。"""
    code: str                                          # 色号，如 '02'
    color_cn: str                                      # 颜色（中文）
    dims: "OrderedDict[tuple, int]" = field(default_factory=OrderedDict)  # (长,宽,高)→箱数

    @property
    def total_boxes(self) -> int:
        return sum(self.dims.values())


@dataclass
class Shipment:
    """一个 PO（PL 内一张有效表）汇总成的一组出货明细。"""
    po: str           # 合同号（含 -N 后缀），来自 E3
    style: str        # 客款号，来自 E5
    fac_style: str    # 厂款号，来自文件名
    dest_cn: str      # 中文目的地，来自 G3
    dest_raw: str     # 原始英文目的地，来自 G2
    ship_date: str = ''   # 出货日期（如 "6.5"），来自 Z3
    src_file: str = ''
    comp: str = ''        # 成分（来自 PL _meta 隐藏表）
    pack: str = ''        # 包装方式（来自 PL _meta 隐藏表）
    tiers: list[Tier] = field(default_factory=list)
    color_rows: list[ColorBoxes] = field(default_factory=list)  # 点箱数明细用

    @property
    def qty(self) -> int:
        return sum(t.pieces for t in self.tiers)

    @property
    def total_boxes(self) -> int:
        return sum(t.boxes for t in self.tiers)

    @property
    def total_net(self) -> float:
        return round(sum(t.net for t in self.tiers), 2)

    @property
    def total_gross(self) -> float:
        return round(sum(t.gross for t in self.tiers), 2)


@dataclass
class GroupSpec:
    """运行时定义的一个出货表分组（对应一批目的地；按日期自动拆文件）。"""
    label: str                    # 表头 H1 文本，如 'BASH 巴安斯&SOUTH AFRICA 南非'
    destinations: list[str]       # 归入本组的中文目的地列表（G3）
    hanger_note: str = DEFAULT_HANGER_NOTE
    dest_tags: dict | None = None  # 中文目的地 → A 列短标签（合并文件时用）


# ── 日期工具 ────────────────────────────────────────────────────────
def _parse_md(s: str):
    """'M.D' / 'M-D' → (month, day) tuple, or None."""
    m = re.match(r'^(\d{1,2})[.\-](\d{1,2})$', (s or '').strip())
    return (int(m.group(1)), int(m.group(2))) if m else None


def _arrive_date(ship_date: str) -> str:
    """factory +1 day; '6.5' → '6.6'; returns original string if unparseable."""
    parsed = _parse_md(ship_date)
    if not parsed:
        return ship_date
    month, day = parsed
    try:
        d = date(date.today().year, month, day) + timedelta(days=1)
        return f'{d.month}.{d.day}'
    except ValueError:
        return ship_date


def _factory_date_str(ship_date: str) -> str:
    """到深圳日期 − 1 day → 出厂日期; '6.6' → '6.5'."""
    parsed = _parse_md(ship_date)
    if not parsed:
        return ship_date
    month, day = parsed
    try:
        d = date(date.today().year, month, day) - timedelta(days=1)
        return f'{d.month}.{d.day}'
    except ValueError:
        return ship_date


def _weekday_str(ship_date: str) -> str:
    """'6.5' → '周五' etc., or '' if unparseable."""
    parsed = _parse_md(ship_date)
    if not parsed:
        return ''
    month, day = parsed
    try:
        d = date(date.today().year, month, day)
        return WEEKDAYS[d.weekday()]
    except ValueError:
        return ''


def _date_sort_key(ship_date: str):
    """Sort key: parse to (month, day) tuple for chronological ordering."""
    parsed = _parse_md(ship_date)
    return parsed if parsed else (99, 99)


# ── PL 解析 ────────────────────────────────────────────────────────
def _num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip().replace(',', '')
    if not s:
        return None
    try:
        return float(s) if '.' in s else int(s)
    except ValueError:
        return None


def _int(v, default=None):
    n = _num(v)
    return int(n) if n is not None else default


def parse_fac_style(fname: str, po: str, style: str) -> str:
    """从文件名 'PO号 客款号 厂款号 PL-目的地.xlsx' 取厂款号。"""
    stem = Path(fname).stem
    m = re.split(r'\s+PL\b', stem, maxsplit=1)
    head = m[0]
    po_base = re.sub(r'-\d+$', '', po)
    prefix = f'{po_base} {style} '
    if head.startswith(prefix):
        return head[len(prefix):].strip()
    parts = head.split()
    return parts[-1] if parts else ''


def _is_valid_sheet(ws, po_base: str) -> bool:
    """有效表：E3 合同号前缀 == 文件名 PO 号，且 E5 款号非空。"""
    e3 = ws.cell(*PL['合同号']).value
    e5 = ws.cell(*PL['客款号']).value
    if not e3 or not e5:
        return False
    return str(e3).strip().startswith(po_base)


def _parse_color_map(ws) -> dict:
    """读 PL 底部色号对照表，返回 颜色名(大写) → (色号, 主布颜色中文)。
    中文名(主布颜色)与英文名(客色号)都映射到同一色号。"""
    cmap: dict[str, tuple] = {}
    for r in range(PL_BOX_FIRST_ROW, ws.max_row + 1):
        if str(ws.cell(r, PL_CMAP_C_CN).value or '').strip() != PL_CMAP_HDR:
            continue
        rr = r + 1
        while rr <= ws.max_row:
            cn = str(ws.cell(rr, PL_CMAP_C_CN).value or '').strip()
            en = str(ws.cell(rr, PL_CMAP_C_EN).value or '').strip()
            code = str(ws.cell(rr, PL_CMAP_C_CODE).value or '').strip()
            if not cn and not en and not code:
                break
            if code:
                if cn:
                    cmap.setdefault(cn.upper(), (code, cn))
                if en:
                    cmap.setdefault(en.upper(), (code, cn))
            rr += 1
        # 对照表可能写两份（装箱单数量/订单原数），第一份即可
        break
    return cmap


def _build_color_rows(color_boxes, cmap) -> list[ColorBoxes]:
    """把 (颜色名→{箱规:箱数}) 按色号归并。中英文同色号合并；
    色号首次出现顺序保留。无对照表时退化为：色号留空、颜色取原名。"""
    by_code: "OrderedDict[str, ColorBoxes]" = OrderedDict()
    for name, dims in color_boxes.items():
        code, color_cn = cmap.get(name.upper(), ('', name))
        ckey = code or f'__{name.upper()}'   # 无色号时按名分组
        cb = by_code.get(ckey)
        if cb is None:
            cb = ColorBoxes(code=code, color_cn=color_cn, dims=OrderedDict())
            by_code[ckey] = cb
        for dim, n in dims.items():
            cb.dims[dim] = cb.dims.get(dim, 0) + n
    return list(by_code.values())


def _read_meta_sheet(wb) -> dict:
    """Read hidden _meta sheet → {厂款号: {'comp': str, 'pack': str}}."""
    try:
        ws = wb['_meta']
    except KeyError:
        return {}
    result = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        style = str(row[0]).strip()
        comp = str(row[1] or '').strip() if len(row) > 1 else ''
        pack = str(row[2] or '').strip() if len(row) > 2 else ''
        if style:
            result[style] = {'comp': comp, 'pack': pack}
    return result


def parse_pl_file(path: Path) -> list[Shipment]:
    """解析一个 PL 文件，返回其中每张有效表对应的 Shipment。"""
    wb = load_workbook(path, data_only=False, read_only=False)
    meta = _read_meta_sheet(wb)
    shipments: list[Shipment] = []
    fname = path.name
    po_base = None
    m = re.match(r'(PO\d+)', fname)
    if m:
        po_base = m.group(1)

    for ws in wb.worksheets:
        if po_base and not _is_valid_sheet(ws, po_base):
            continue
        e3 = str(ws.cell(*PL['合同号']).value or '').strip()
        e5 = str(ws.cell(*PL['客款号']).value or '').strip()
        if not e3 or not e5:
            continue
        dest_cn = str(ws.cell(*PL['目的地中']).value or '').strip()
        dest_raw = str(ws.cell(*PL['目的地英']).value or '').strip()
        nw = _num(ws.cell(*PL['NW']).value) or 0.0
        ship_date = str(ws.cell(*PL['出货日期']).value or '').strip()
        fac = parse_fac_style(fname, e3, e5)

        cmap = _parse_color_map(ws)

        tiers_by_dim: OrderedDict[tuple, Tier] = OrderedDict()
        # 点箱数明细用：按 (颜色名, 箱规) 记箱数，保留颜色名首次出现顺序
        color_boxes: "OrderedDict[str, OrderedDict[tuple, int]]" = OrderedDict()
        cur_color = ''   # 颜色块首行写色名，续行 A 列为空 → 沿用上一个色名
        for r in range(PL_BOX_FIRST_ROW, ws.max_row + 1):
            d_end = ws.cell(r, PL_C_BOXNO_END).value
            if isinstance(d_end, str) and d_end.strip() == '共':
                break
            b = _int(ws.cell(r, PL_C_BOXNO_START).value)
            de = _int(d_end)
            length = _int(ws.cell(r, PL_C_LEN).value)
            if b is None or de is None or length is None:
                continue
            width = _int(ws.cell(r, PL_C_WID).value) or 0
            height = _int(ws.cell(r, PL_C_HGT).value) or 0
            carton = _num(ws.cell(r, PL_C_CARTON).value) or 0.0
            count = de - b + 1
            if count <= 0:
                continue
            pieces_per_box = 0
            for c in range(PL_C_SIZE_FIRST, PL_C_SIZE_LAST + 1):
                v = _num(ws.cell(r, c).value)
                if isinstance(v, (int, float)):
                    pieces_per_box += v
            pieces_per_box = int(round(pieces_per_box))

            key = (length, width, height)
            t = tiers_by_dim.get(key)
            if t is None:
                t = Tier(length, width, height)
                tiers_by_dim[key] = t
            t.boxes += count
            t.pieces += pieces_per_box * count
            net_box = pieces_per_box * nw
            t.net += net_box * count
            t.gross += (net_box + carton) * count
            t.cbm += (length * width * height / 1_000_000.0) * count

            color_name = str(ws.cell(r, PL_C_COLOR).value or '').strip()
            if color_name:
                cur_color = color_name
            if cur_color:
                cb = color_boxes.setdefault(cur_color, OrderedDict())
                cb[key] = cb.get(key, 0) + count

        if not tiers_by_dim:
            continue
        tiers = sorted(tiers_by_dim.values(), key=lambda t: t.height, reverse=True)
        for t in tiers:
            t.net = round(t.net, 2)
            t.gross = round(t.gross, 2)
            t.cbm = round(t.cbm, 6)
        color_rows = _build_color_rows(color_boxes, cmap)
        _meta_e = meta.get(fac, {})
        shipments.append(Shipment(
            po=e3, style=e5, fac_style=fac, dest_cn=dest_cn, dest_raw=dest_raw,
            ship_date=ship_date, src_file=fname,
            comp=_meta_e.get('comp', ''), pack=_meta_e.get('pack', ''),
            tiers=tiers, color_rows=color_rows))
    wb.close()
    return shipments


def parse_pl_folder(folder) -> OrderedDict[str, list[Shipment]]:
    """扫描文件夹内全部 PL，返回 中文目的地 → [Shipment]（按文件名顺序）。"""
    folder = Path(folder)
    by_dest: OrderedDict[str, list[Shipment]] = OrderedDict()
    for f in sorted(folder.glob('*.xlsx')):
        if f.name.startswith('~$'):
            continue
        try:
            for sh in parse_pl_file(f):
                by_dest.setdefault(sh.dest_cn, []).append(sh)
        except Exception as e:                          # noqa: BLE001
            print(f'  [跳过] {f.name}: {e}')
    return by_dest


# ── 样式 ────────────────────────────────────────────────────────────
_thin = Side(style='thin')
BORDER = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
FONT = Font(name='宋体', size=14)
FONT_CBM = Font(name='Times New Roman', size=14)
FONT_TITLE = Font(name='宋体', size=20)
FONT_16 = Font(name='宋体', size=16)
FONT_16B = Font(name='宋体', size=16, bold=True)
FONT_15B = Font(name='宋体', size=15, bold=True)
FONT_12 = Font(name='宋体', size=12)
FONT_12B = Font(name='宋体', size=12, bold=True)
CENTER = Alignment(horizontal='center', vertical='center')
CENTER_WRAP = Alignment(horizontal='center', vertical='center', wrap_text=True)
LEFT = Alignment(horizontal='left', vertical='center')
VCENTER = Alignment(vertical='center')
FILL_YELLOW = PatternFill('solid', fgColor='FFFF00')
FILL_CYAN = PatternFill('solid', fgColor='00FFFF')

NF_INT = '0_ '
NF_DIM = '#,##0_ '
NF_CBM = '#,##0.000_ '
NF_WT = '0.00_);[Red]\\(0.00\\)'
NF_WT_TOTAL = '0.00_ '

# 列宽：对应新 COL 映射（A-T，含包装方式列）
COL_WIDTHS = {
    'A': 13.8,   # tag
    'B': 17.97,  # 合同号
    'C': 13.0,   # 客款号
    'D': 18.0,   # 厂款号
    'E': 49.85,  # 成分
    'F': 16.0,   # 包装方式
    'G': 10.14,  # 数量
    'H': 8.88,   # 总箱数
    'I': 8.88,   # 箱数
    'J': 6.53,   # 长
    'K': 4.63,   # X
    'L': 6.53,   # 宽
    'M': 4.63,   # X
    'N': 10.5,   # 高（兼放页脚「总立方数」标签）
    'O': 11.72,  # CBM
    'P': 10.47,  # 净重
    'Q': 10.93,  # 毛重
    'R': 12.65,  # 总净重
    'S': 12.65,  # 总毛重
    'T': 19.37,  # 品牌
}


def _set(ws, r, c, value, font=FONT, align=VCENTER, border=None,
         numfmt=None, fill=None):
    cell = ws.cell(r, c, value)
    cell.font = font
    cell.alignment = align
    if border:
        cell.border = border
    if numfmt:
        cell.number_format = numfmt
    if fill:
        cell.fill = fill
    return cell


def _box_border(ws, r1, c1, r2, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = BORDER


# ── 出货表写入 ──────────────────────────────────────────────────────
def _short_tag(dest_cn: str) -> str:
    return dest_cn or ''


def _write_header_block(ws, spec: GroupSpec, opts: dict):
    factory_date = opts.get('factory_date', '')
    arrive_date = opts.get('arrive_date', '')
    weekday = opts.get('weekday', '')
    code = opts.get('code', '')

    _set(ws, 1, COL['合同号'], DEFAULT_MAKER_TO, font=FONT)
    _set(ws, 1, COL['总箱数'], spec.label, font=FONT_TITLE, align=CENTER)
    if weekday:
        _set(ws, 1, COL['净重'], weekday, font=FONT_16,
             align=Alignment(horizontal='right', vertical='center'))
    _set(ws, 1, COL['总毛重'], '编号：', font=FONT_16,
         align=Alignment(horizontal='right', vertical='center'))
    _set(ws, 1, COL['品牌'], code, font=FONT, align=CENTER)

    intro = (f'普宁集杰有货从集杰出货，在{factory_date}出厂,{arrive_date}需到深圳。'
             f'不能早也不能迟，请和厂家联系取货，资料如下：')
    _set(ws, 3, COL['合同号'], intro, font=FONT, align=LEFT)
    _set(ws, 4, COL['合同号'], '请留意，谢谢。', font=FONT, align=LEFT)

    # 「货号是客款号」跨 总箱数..长（H5:J5）
    ws.merge_cells(start_row=5, start_column=COL['总箱数'],
                   end_row=5, end_column=COL['长'])
    _set(ws, 5, COL['总箱数'], '货号是客款号', font=FONT_15B, align=CENTER_WRAP)

    _set(ws, 6, COL['合同号'], opts.get('factory_addr', DEFAULT_FACTORY_ADDR),
         font=FONT, align=LEFT)
    _set(ws, 7, COL['合同号'], opts.get('contact', DEFAULT_CONTACT),
         font=FONT, align=LEFT)
    _set(ws, 8, COL['合同号'], opts.get('phone', DEFAULT_PHONE),
         font=FONT, align=LEFT)

    # 表头行 (row 9)
    headers = [
        ('合同号', '合同号'), ('客款号', '客款号'),
        ('厂款号', '厂款号'), ('成分', '成分'), ('包装方式', '包装方式'), ('数量', '数量'),
        ('总箱数', '总箱数'), ('箱数', '箱数'), ('长', '箱尺寸(两种）'),
        ('CBM', 'CBM'), ('净重', '净重'), ('毛重', '毛重'),
        ('总净重', '总净重'), ('总毛重', '总毛重'), ('品牌', '品牌'),
    ]
    for key, text in headers:
        _set(ws, 9, COL[key], text, font=FONT, align=CENTER, border=BORDER)
    # 箱尺寸表头跨 长..高（J9:N9）
    ws.merge_cells(start_row=9, start_column=COL['长'],
                   end_row=9, end_column=COL['高'])
    _box_border(ws, 9, COL['长'], 9, COL['高'])
    for c in range(COL['合同号'], COL['品牌'] + 1):
        ws.cell(9, c).border = BORDER


def _write_po_rows(ws, start_row: int, shipments: list[Shipment],
                   spec: GroupSpec, multi: bool) -> int:
    r = start_row
    tags = spec.dest_tags or {}
    for sh in shipments:
        n = len(sh.tiers)
        r0 = r
        merge_keys = ['合同号', '客款号', '厂款号', '成分', '包装方式', '数量',
                      '总箱数', '总净重', '总毛重']
        for i, t in enumerate(sh.tiers):
            rr = r0 + i
            ws.row_dimensions[rr].height = 23
            if multi and i == 0:
                _set(ws, rr, COL['tag'], tags.get(sh.dest_cn, _short_tag(sh.dest_cn)),
                     font=FONT, align=CENTER, border=BORDER)
            # 每档独有列
            _set(ws, rr, COL['箱数'], t.boxes, font=FONT, align=CENTER,
                 border=BORDER, numfmt=NF_INT)
            _set(ws, rr, COL['长'], t.length, font=FONT, align=CENTER,
                 border=BORDER, numfmt=NF_DIM)
            _set(ws, rr, COL['X1'], 'X', font=FONT, align=CENTER,
                 border=BORDER, numfmt=NF_DIM)
            _set(ws, rr, COL['宽'], t.width, font=FONT, align=CENTER,
                 border=BORDER, numfmt=NF_DIM)
            _set(ws, rr, COL['X2'], 'X', font=FONT, align=CENTER,
                 border=BORDER, numfmt=NF_DIM)
            _set(ws, rr, COL['高'], t.height, font=FONT, align=CENTER,
                 border=BORDER, numfmt=NF_DIM)
            _set(ws, rr, COL['CBM'], round(t.cbm, 6), font=FONT_CBM,
                 align=CENTER, border=BORDER, numfmt=NF_CBM)
            _set(ws, rr, COL['净重'], t.net, font=FONT, align=CENTER,
                 border=BORDER, numfmt=NF_WT)
            _set(ws, rr, COL['毛重'], t.gross, font=FONT, align=CENTER,
                 border=BORDER, numfmt=NF_WT)
            _set(ws, rr, COL['品牌'], DEFAULT_BRAND, font=FONT,
                 align=CENTER, border=BORDER)


        # 合并列（写首行，多档时向下合并）
        _set(ws, r0, COL['合同号'], sh.po, font=FONT, align=CENTER, border=BORDER)
        _set(ws, r0, COL['客款号'], _num_or_text(sh.style), font=FONT,
             align=CENTER, border=BORDER, numfmt=NF_INT)
        _set(ws, r0, COL['厂款号'], sh.fac_style, font=FONT, align=CENTER, border=BORDER)
        _set(ws, r0, COL['成分'], sh.comp, font=FONT, align=CENTER, border=BORDER)
        _set(ws, r0, COL['包装方式'], sh.pack, font=FONT, align=CENTER, border=BORDER)
        _set(ws, r0, COL['数量'], sh.qty, font=FONT, align=CENTER, border=BORDER,
             numfmt=NF_INT)
        _set(ws, r0, COL['总箱数'], sh.total_boxes, font=FONT, align=CENTER,
             border=BORDER, numfmt=NF_INT)
        _set(ws, r0, COL['总净重'], sh.total_net, font=FONT, align=CENTER,
             border=BORDER, numfmt=NF_WT)
        _set(ws, r0, COL['总毛重'], sh.total_gross, font=FONT, align=CENTER,
             border=BORDER, numfmt=NF_WT)
        if n > 1:
            for key in merge_keys:
                ws.merge_cells(start_row=r0, start_column=COL[key],
                               end_row=r0 + n - 1, end_column=COL[key])
            _box_border(ws, r0, COL['合同号'], r0 + n - 1, COL['品牌'])
        r += n
    return r


def _write_totals(ws, r: int, shipments: list[Shipment]) -> int:
    qty = sum(s.qty for s in shipments)
    boxes = sum(s.total_boxes for s in shipments)
    net = round(sum(s.total_net for s in shipments), 2)
    gross = round(sum(s.total_gross for s in shipments), 2)
    cbm = round(sum(t.cbm for s in shipments for t in s.tiers), 6)
    ws.row_dimensions[r].height = 25
    _set(ws, r, COL['成分'], '总计:', font=FONT, align=LEFT)
    _set(ws, r, COL['数量'], qty, font=FONT, align=CENTER, numfmt=NF_INT)
    _set(ws, r, COL['总箱数'], boxes, font=FONT, align=CENTER, numfmt=NF_INT)
    _set(ws, r, COL['长'], '箱', font=FONT, align=CENTER, numfmt=NF_DIM)
    _set(ws, r, COL['高'], '共', font=FONT, align=CENTER, numfmt=NF_DIM)
    _set(ws, r, COL['CBM'], cbm, font=FONT, align=CENTER, numfmt=NF_CBM)
    _set(ws, r, COL['总净重'], net, font=FONT, align=CENTER, numfmt=NF_WT_TOTAL)
    _set(ws, r, COL['总毛重'], gross, font=FONT, align=CENTER, numfmt=NF_WT_TOTAL)
    _set(ws, r + 1, COL['数量'], '件', font=FONT, align=CENTER, numfmt=NF_INT)
    return r + 2


def _write_footer(ws, r: int, shipments: list[Shipment], spec: GroupSpec):
    r += 1
    # ***报关*** 行兼 实际量到的箱尺寸 表头行
    _set(ws, r, COL['合同号'], DEFAULT_REPORT_NOTE, font=FONT_12, align=LEFT)
    _set(ws, r, COL['箱数'], '箱数', font=FONT_12B, align=CENTER, border=BORDER)
    ws.merge_cells(start_row=r, start_column=COL['长'], end_row=r,
                   end_column=COL['高'])
    _set(ws, r, COL['长'], '实际量到的箱尺寸', font=FONT_12B, align=CENTER)
    _box_border(ws, r, COL['长'], r, COL['高'])
    _set(ws, r, COL['CBM'], '实际CBM', font=FONT_12B, align=CENTER, border=BORDER)
    r += 1
    # 汇总全表箱规
    agg: OrderedDict[tuple, list] = OrderedDict()
    for s in shipments:
        for t in s.tiers:
            key = (t.length, t.width, t.height)
            cur = agg.setdefault(key, [0, 0.0])
            cur[0] += t.boxes
            cur[1] += t.cbm
    tiers = sorted(agg.items(), key=lambda kv: kv[0][2], reverse=True)
    first_actual = r
    for (length, width, height), (boxes, cbm) in tiers:
        _set(ws, r, COL['箱数'], boxes, font=FONT, align=CENTER, border=BORDER,
             numfmt=NF_INT)
        _set(ws, r, COL['长'], length, font=FONT, align=CENTER, border=BORDER,
             numfmt=NF_DIM, fill=FILL_CYAN)
        _set(ws, r, COL['X1'], 'X', font=FONT, align=CENTER, border=BORDER,
             numfmt=NF_DIM, fill=FILL_CYAN)
        _set(ws, r, COL['宽'], width, font=FONT, align=CENTER, border=BORDER,
             numfmt=NF_DIM, fill=FILL_CYAN)
        _set(ws, r, COL['X2'], 'X', font=FONT, align=CENTER, border=BORDER,
             numfmt=NF_DIM, fill=FILL_CYAN)
        _set(ws, r, COL['高'], height, font=FONT, align=CENTER, border=BORDER,
             numfmt=NF_DIM, fill=FILL_CYAN)
        _set(ws, r, COL['CBM'], round(cbm, 6), font=FONT_CBM, align=CENTER,
             border=BORDER, numfmt=NF_CBM, fill=FILL_CYAN)
        r += 1
    ws.cell(first_actual, COL['合同号'], DEFAULT_QUESTION_NOTE).font = FONT_12
    ws.cell(first_actual, COL['合同号']).alignment = LEFT
    total_boxes = sum(v[0] for v in agg.values())
    total_cbm = round(sum(v[1] for v in agg.values()), 6)
    _set(ws, r, COL['总箱数'], '总箱数', font=FONT_12B, align=CENTER, border=BORDER)
    _set(ws, r, COL['箱数'], total_boxes, font=FONT_12B, align=CENTER,
         border=BORDER, numfmt=NF_INT)
    _set(ws, r, COL['高'], '总立方数', font=FONT_12B, align=CENTER, border=BORDER)
    _set(ws, r, COL['CBM'], total_cbm, font=FONT_CBM, align=CENTER, border=BORDER,
         numfmt=NF_CBM)
    return r


def _num_or_text(s: str):
    n = _num(s)
    return n if n is not None else s


def _apply_dimensions(ws):
    for col, w in COL_WIDTHS.items():
        ws.column_dimensions[col].width = w
    ws.row_dimensions[1].height = 25
    for rr in range(2, 9):
        ws.row_dimensions[rr].height = 20
    ws.row_dimensions[9].height = 26


def build_workbook(spec: GroupSpec, shipments: list[Shipment],
                   opts: dict, multi: bool) -> Workbook:
    wb = Workbook()
    ws = wb.active
    ws.title = 'Sheet1'
    _write_header_block(ws, spec, opts)
    next_row = _write_po_rows(ws, 10, shipments, spec, multi)
    next_row = _write_totals(ws, next_row, shipments)
    _write_footer(ws, next_row, shipments, spec)
    _apply_dimensions(ws)
    return wb


# ── 顶层导出 ────────────────────────────────────────────────────────
def export_shipping_tables(by_dest: OrderedDict[str, list[Shipment]],
                           out_dir,
                           groups: list[GroupSpec] | None = None,
                           code_date_str: str = '',
                           code_start: int = 1) -> list[str]:
    """按「分组 × 出货日期」导出出货表。

    by_dest        : parse_pl_folder 的结果。
    groups         : 运行时分组；None 时每个目的地各一组。
    code_date_str  : 编号日期前缀 (YYYYMMDD)，默认取今天。
    code_start     : 编号起始序号（如 1 → 01）。
    返回写出的文件名列表。
    """
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)

    if not code_date_str:
        code_date_str = date.today().strftime('%Y%m%d')

    if groups is None:
        groups = [GroupSpec(label=dest, destinations=[dest])
                  for dest in by_dest.keys()]

    written = []
    code_idx = code_start
    for g in groups:
        all_ships = [s for d in g.destinations for s in by_dest.get(d, [])]
        if not all_ships:
            print(f'  [空组] {g.label}：无匹配目的地，跳过')
            continue

        # 按出货日期分组，空日期归入 '(未填日期)'，按日期排序
        by_date: OrderedDict[str, list[Shipment]] = OrderedDict()
        for sh in all_ships:
            key = sh.ship_date or '(未填日期)'
            by_date.setdefault(key, []).append(sh)
        ordered_dates = sorted(by_date.keys(), key=_date_sort_key)

        for dk in ordered_dates:
            date_ships = by_date[dk]
            spec_dk, multi = _subset_spec(g, date_ships)
            code = f'CO{code_date_str}{code_idx:02d}'
            code_idx += 1
            opts = {
                'factory_date': _factory_date_str(dk) if dk != '(未填日期)' else '',
                'arrive_date':  dk                    if dk != '(未填日期)' else '',
                'weekday':      _weekday_str(_factory_date_str(dk)) if dk != '(未填日期)' else '',
                'code': code,
            }
            wb = build_workbook(spec_dk, date_ships, opts, multi)
            date_tag = dk.replace('.', '-').replace('(未填日期)', '')
            fname = _safe_name(f'出货表 {date_tag} 集杰 {spec_dk.label}.xlsx')
            wb.save(out / fname)
            written.append(fname)
            print(f'  [写出] {fname}  ({len(date_ships)} 个PO，'
                  f'{sum(s.total_boxes for s in date_ships)} 箱，编号 {code})')
    return written


def _safe_name(name: str) -> str:
    name = re.sub(r'[\\/:*?"<>|]', '-', name)
    return re.sub(r'\s+', ' ', name).strip()


def _subset_spec(g: GroupSpec, ships: list[Shipment]) -> tuple[GroupSpec, bool]:
    """返回仅含 ships 中实际出现的目的地的 GroupSpec（保留分组内顺序）。

    合并分组里某个出货日期可能只有部分目的地有货 —— 此时文件名/标题
    应只体现当天有货的目的地，而非整个合并分组。
    若全部目的地都有货，则原样返回 g（保留用户可能自定义的标签）。
    """
    present = [d for d in g.destinations
               if any(s.dest_cn == d for s in ships)]
    if not present or len(present) == len(g.destinations):
        return g, len(g.destinations) > 1
    sub = GroupSpec(label=' + '.join(present), destinations=present,
                    hanger_note=g.hanger_note, dest_tags=g.dest_tags)
    return sub, len(present) > 1


# ── 点箱数明细 写入 ─────────────────────────────────────────────────
# 列：A 子目的地标签 | B 合同号 | C 款号 | D 外箱序号 | E 色号 | F 颜色
#     | (各箱规箱数...) | 总数
DK_C_TAG = 1
DK_C_PO = 2
DK_C_STYLE = 3
DK_C_SEQ = 4
DK_C_CODE = 5
DK_C_COLOR = 6
DK_C_DIM0 = 7            # 第一个箱规列
DK_NOTE_FONT = Font(name='宋体', size=11)
DK_TITLE_FONT = Font(name='宋体', size=16, bold=True)
DK_HDR_FONT = Font(name='宋体', size=12, bold=True)
LEFT_WRAP = Alignment(horizontal='left', vertical='center', wrap_text=True)
DK_DATE_FMT = 'mm-dd-yy'


def _dianka_dim_order(shipments: list[Shipment]) -> list[tuple]:
    """组内全部箱规，按首次出现顺序去重。"""
    order: list[tuple] = []
    seen = set()
    for sh in shipments:
        for cr in sh.color_rows:
            for dim in cr.dims:
                if dim not in seen:
                    seen.add(dim)
                    order.append(dim)
    return order


def _write_dianka_sheet(ws, spec: GroupSpec, shipments: list[Shipment],
                        ship_date: str, multi: bool):
    dims = _dianka_dim_order(shipments)
    n_dim = len(dims)
    c_total = DK_C_DIM0 + n_dim          # 总数列
    last_col = c_total

    # 顶部说明
    ws.merge_cells(start_row=1, start_column=DK_C_PO, end_row=1, end_column=last_col)
    _set(ws, 1, DK_C_PO, DIANKA_NOTE1, font=DK_NOTE_FONT, align=LEFT_WRAP)
    ws.merge_cells(start_row=2, start_column=DK_C_PO, end_row=2, end_column=last_col)
    _set(ws, 2, DK_C_PO, DIANKA_NOTE2, font=DK_NOTE_FONT, align=LEFT_WRAP)
    ws.row_dimensions[2].height = 60
    ws.merge_cells(start_row=3, start_column=DK_C_PO, end_row=3, end_column=last_col)
    _dk_styles = sorted({str(_num_or_text(sh.style)) for sh in shipments if sh.style})
    _dk_title = (f'{spec.label} {"+".join(_dk_styles)}走货通知'
                 if _dk_styles else spec.label)
    _set(ws, 3, DK_C_PO, _dk_title, font=DK_TITLE_FONT, align=CENTER)

    # 表头（第 4 行）
    _set(ws, 4, DK_C_PO, '合同号', font=DK_HDR_FONT, align=CENTER, border=BORDER)
    _set(ws, 4, DK_C_STYLE, '款号', font=DK_HDR_FONT, align=CENTER, border=BORDER)
    _set(ws, 4, DK_C_SEQ, '外箱序号', font=DK_HDR_FONT, align=CENTER, border=BORDER)
    _set(ws, 4, DK_C_CODE, '色号', font=DK_HDR_FONT, align=CENTER, border=BORDER)
    _set(ws, 4, DK_C_COLOR, '颜色', font=DK_HDR_FONT, align=CENTER, border=BORDER)
    for i, (L, W, H) in enumerate(dims):
        _set(ws, 4, DK_C_DIM0 + i, f'{L}*{W}*{H}/\n箱数',
             font=DK_HDR_FONT, align=CENTER_WRAP, border=BORDER)
    _set(ws, 4, c_total, '总数', font=DK_HDR_FONT, align=CENTER, border=BORDER)
    if multi:
        ws.cell(4, DK_C_TAG).border = BORDER

    # 数据行
    r = 5
    seq = 1
    dim_idx = {d: DK_C_DIM0 + i for i, d in enumerate(dims)}
    for sh in shipments:
        rows = sh.color_rows or []
        if not rows:
            continue
        block_fr = r
        for cr in rows:
            if multi:
                if cr is rows[0]:
                    _set(ws, r, DK_C_TAG, _short_tag(sh.dest_cn),
                         font=FONT, align=CENTER, border=BORDER)
                else:
                    ws.cell(r, DK_C_TAG).border = BORDER
            _set(ws, r, DK_C_SEQ, seq, font=FONT, align=CENTER,
                 border=BORDER, numfmt=NF_INT)
            _set(ws, r, DK_C_CODE, cr.code, font=FONT, align=CENTER, border=BORDER)
            _set(ws, r, DK_C_COLOR, cr.color_cn, font=FONT, align=CENTER, border=BORDER)
            for d, c in dim_idx.items():
                n = cr.dims.get(d)
                _set(ws, r, c, n if n else None, font=FONT, align=CENTER,
                     border=BORDER, numfmt=NF_INT)
            _set(ws, r, c_total, cr.total_boxes, font=FONT, align=CENTER,
                 border=BORDER, numfmt=NF_INT)
            seq += 1
            r += 1
        block_lr = r - 1
        # 合同号 / 款号 跨该 PO 各色号行合并
        _set(ws, block_fr, DK_C_PO, sh.po, font=FONT, align=CENTER, border=BORDER)
        _set(ws, block_fr, DK_C_STYLE, _num_or_text(sh.style), font=FONT,
             align=CENTER, border=BORDER, numfmt=NF_INT)
        if block_lr > block_fr:
            ws.merge_cells(start_row=block_fr, start_column=DK_C_PO,
                           end_row=block_lr, end_column=DK_C_PO)
            ws.merge_cells(start_row=block_fr, start_column=DK_C_STYLE,
                           end_row=block_lr, end_column=DK_C_STYLE)
            for rr in range(block_fr + 1, block_lr + 1):
                ws.cell(rr, DK_C_PO).border = BORDER
                ws.cell(rr, DK_C_STYLE).border = BORDER

    last_data = r - 1
    # 合计行：各箱规列求和 + 总数总计
    r += 1
    total_row = r
    grand = 0
    for d, c in dim_idx.items():
        col_sum = sum(cr.dims.get(d, 0) for sh in shipments for cr in sh.color_rows)
        grand += col_sum
        _set(ws, total_row, c, col_sum or None, font=DK_HDR_FONT, align=CENTER,
             border=BORDER, numfmt=NF_INT)
    _set(ws, total_row, DK_C_COLOR, '合计', font=DK_HDR_FONT, align=CENTER, border=BORDER)
    _set(ws, total_row, c_total, grand, font=DK_HDR_FONT, align=CENTER,
         border=BORDER, numfmt=NF_INT)

    # 日期（合计下方两行，右侧）— 写入文件生成当天日期
    dcell = ws.cell(total_row + 2, c_total, date.today())
    dcell.number_format = DK_DATE_FMT
    dcell.font = FONT
    dcell.alignment = Alignment(horizontal='right', vertical='center')

    # 列宽
    widths = {DK_C_TAG: 9.7, DK_C_PO: 17.7, DK_C_STYLE: 17.0, DK_C_SEQ: 14.3,
              DK_C_CODE: 12.3, DK_C_COLOR: 24.6}
    for c, w in widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    for i in range(n_dim):
        ws.column_dimensions[get_column_letter(DK_C_DIM0 + i)].width = 17.0
    ws.column_dimensions[get_column_letter(c_total)].width = 22.2
    return last_data


def _write_dianka_summary(ws, groups_here: list[tuple]) -> None:
    """汇总表：每个分组一行，汇总件数/箱数/CBM。"""
    SUM_HDR = Font(name='宋体', size=12, bold=True)
    SUM_FONT = Font(name='宋体', size=11)

    # Row 2: combined styles title
    all_styles = sorted({str(_num_or_text(sh.style))
                         for g, ships, _ in groups_here
                         for sh in ships if sh.style})
    ws.cell(2, 2, '+'.join(all_styles)).font = SUM_HDR

    # Row 3: headers
    for col, hdr in ((2, '款号'), (3, '客户'), (4, '件数'),
                     (5, '箱数'), (6, '体积方数'), (7, '装箱作业员')):
        c = ws.cell(3, col, hdr)
        c.font = SUM_HDR
        c.border = BORDER
        c.alignment = CENTER

    # Data rows
    r = 4
    tot_qty = tot_boxes = tot_cbm = 0
    for g, ships, _ in groups_here:
        grp_styles = sorted({str(_num_or_text(sh.style)) for sh in ships if sh.style})
        qty   = sum(sh.qty for sh in ships)
        boxes = sum(sh.total_boxes for sh in ships)
        cbm   = round(sum(t.cbm for sh in ships for t in sh.tiers), 2)
        tot_qty   += qty
        tot_boxes += boxes
        tot_cbm    = round(tot_cbm + cbm, 2)
        for col, val in ((2, '+'.join(grp_styles)), (3, g.label),
                         (4, qty), (5, boxes), (6, cbm), (7, '')):
            c = ws.cell(r, col, val)
            c.font = SUM_FONT
            c.border = BORDER
            c.alignment = CENTER
        r += 1

    # 合计 row
    r += 1
    for col, val in ((3, '合计'), (4, tot_qty), (5, tot_boxes), (6, tot_cbm)):
        c = ws.cell(r, col, val)
        c.font = SUM_HDR
        c.border = BORDER
        c.alignment = CENTER

    # Column widths
    for col_letter, w in (('B', 18), ('C', 26), ('D', 10),
                           ('E', 10), ('F', 12), ('G', 18)):
        ws.column_dimensions[col_letter].width = w


def export_dianka_workbooks(by_dest: OrderedDict[str, list[Shipment]],
                            out_dir,
                            groups: list[GroupSpec] | None = None) -> list[str]:
    """按出货日期导出「点箱数明细」：每个日期一个工作簿，每个分组一张表。
    数据来源与出货表相同（PL）。返回写出的文件名列表。"""
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)

    if groups is None:
        groups = [GroupSpec(label=dest, destinations=[dest])
                  for dest in by_dest.keys()]

    # 透视：日期 → [(spec, 该日期的 shipments)]
    by_date: OrderedDict[str, list[tuple]] = OrderedDict()
    for g in groups:
        all_ships = [s for d in g.destinations for s in by_dest.get(d, [])]
        if not all_ships:
            continue
        date_ships: OrderedDict[str, list[Shipment]] = OrderedDict()
        for sh in all_ships:
            date_ships.setdefault(sh.ship_date or '(未填日期)', []).append(sh)
        for dk, ships in date_ships.items():
            spec_dk, multi = _subset_spec(g, ships)
            by_date.setdefault(dk, []).append((spec_dk, ships, multi))

    written = []
    for dk in sorted(by_date.keys(), key=_date_sort_key):
        groups_here = by_date[dk]
        wb = Workbook()
        wb.remove(wb.active)
        used_titles = set()
        for g, ships, multi in groups_here:
            title = _safe_sheet_title(g.label, used_titles)
            ws = wb.create_sheet(title=title)
            ship_date = dk if dk != '(未填日期)' else ''
            _write_dianka_sheet(ws, g, ships, ship_date, multi)
        ws_sum = wb.create_sheet(title='汇总')
        _write_dianka_summary(ws_sum, groups_here)
        date_tag = dk.replace('.', '-').replace('(未填日期)', '无日期')
        fname = _safe_name(f'点箱数明细 {date_tag}.xlsx')
        wb.save(out / fname)
        written.append(fname)
        print(f'  [写出] {fname}  （{len(groups_here)} 张表）')
    return written


def _safe_sheet_title(label: str, used: set) -> str:
    """Excel 表名 ≤31 字符，不含 []:*?/\\，且不重复。"""
    title = re.sub(r'[\[\]:*?/\\]', ' ', label).strip()[:31] or 'Sheet'
    base = title
    i = 2
    while title in used:
        suffix = f' {i}'
        title = base[:31 - len(suffix)] + suffix
        i += 1
    used.add(title)
    return title


# ── CLI ─────────────────────────────────────────────────────────────
def _main(argv):
    if len(argv) < 2:
        print('用法: python shipping_table_export.py <PL文件夹> <输出文件夹> '
              '[--groups groups.json] [--code-date YYYYMMDD] [--code-start N]')
        return 1
    pl_folder, out_dir = argv[0], argv[1]
    groups = None
    code_date_str = ''
    code_start = 1
    if '--groups' in argv:
        gpath = argv[argv.index('--groups') + 1]
        data = json.loads(Path(gpath).read_text(encoding='utf-8'))
        glist = data.get('groups', data) if isinstance(data, dict) else data
        groups = [GroupSpec(label=g['label'], destinations=g['destinations'],
                            hanger_note=g.get('hanger_note', DEFAULT_HANGER_NOTE),
                            dest_tags=g.get('dest_tags'))
                  for g in glist]
    if '--code-date' in argv:
        code_date_str = argv[argv.index('--code-date') + 1]
    if '--code-start' in argv:
        code_start = int(argv[argv.index('--code-start') + 1])

    print(f'扫描 PL 文件夹: {pl_folder}')
    by_dest = parse_pl_folder(pl_folder)
    print(f'发现目的地 {len(by_dest)} 个:')
    for dest, lst in by_dest.items():
        dates = sorted({s.ship_date for s in lst if s.ship_date}, key=_date_sort_key)
        boxes = sum(s.total_boxes for s in lst)
        print(f'  - {dest}: {len(lst)} 个PO，{boxes} 箱，日期: {dates or ["(未填)"]}'
              )
    written = export_shipping_tables(by_dest, out_dir, groups,
                                     code_date_str, code_start)
    print(f'完成，共写出 {len(written)} 个文件 → {out_dir}')
    return 0


if __name__ == '__main__':
    raise SystemExit(_main(sys.argv[1:]))
