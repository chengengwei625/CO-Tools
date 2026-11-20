#!/usr/bin/env python3
"""跟单表生成工具 — 从 港口明细汇总 / 数量明细汇总 生成其余工作表。"""

import json
import csv
import math
import re
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
from copy import copy
from collections import OrderedDict, defaultdict
from datetime import date, datetime

import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.pagebreak import Break
from openpyxl.worksheet.properties import PageSetupProperties

# ══════════════════════════════════════════════════════════════
# 可配置的港口 → 工作表映射（每次根据需要修改下方字典）
# ══════════════════════════════════════════════════════════════

PORT_TAB_CONFIG = {
    '巴西': {
        'ports': ['巴西'],
        'notes': (
            '注:\n'
            '1：巴西港口,尺码牌不同 ,注意与大货区分出来,回货单上面注明"巴西港口"，并且一色齐码齐数后查好再送回\n'
            '2：此单一式两份，放一份在货里随货送回公司\n'
            '3：特殊港口的货，请工厂物必重视，按要求回货，否则一次警告，二次罚款200元/次'
        ),
    },
    '英国诺丁汉': {
        'ports': ['英国诺丁汉'],
        'notes': (
            '注:\n'
            '1：特殊港口,洗水唛上面,有 * ,注意与大货区分出来,回货单上面注明"{port}"，并且一色齐码齐数后查好再送回\n'
            '2：此单放在货里随货送回公司\n'
            '3：特殊港口的货，请工厂物必重视，按要求回货，否则一次警告，二次罚款200元/次'
        ),
    },
    '特殊港口': {
        'ports': ['英国伦敦', 'ASOS UK 英国', 'ASOS EU 欧洲'],
        'notes': (
            '注:\n'
            '1：特殊港口,洗水唛上面,有 * ,注意与大货区分出来,回货单上面注明"{port}"，并且一色齐码齐数后查好再送回\n'
            '2：此单放在货里随货送回公司\n'
            '3：特殊港口的货，请工厂物必重视，按要求回货，否则一次警告，二次罚款200元/次'
        ),
    },
}

# Directory of the running program. For a PyInstaller --onefile exe, __file__
# points to a temp extraction folder that is deleted on exit, so use the exe's
# real location instead (so tab_gen_config.json persists next to the exe).
if getattr(sys, 'frozen', False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).parent

CONFIG_FILE = APP_DIR / 'tab_gen_config.json'

# ══════════════════════════════════════════════════════════════
# 港口配置模板
# ══════════════════════════════════════════════════════════════

def create_port_template(path):
    """Create a port-config template with one sheet per special-port tab."""
    wb = Workbook()
    wb.remove(wb.active)
    for tab_name, cfg in PORT_TAB_CONFIG.items():
        ws = wb.create_sheet(title=tab_name)
        ws.column_dimensions['A'].width = 30
        hdr = ws.cell(row=1, column=1, value='港口名称')
        hdr.font      = Font(size=13, bold=True)
        hdr.alignment = Alignment(horizontal='center', vertical='center')
        hdr.fill      = PatternFill('solid', fgColor='BDD7EE')
        ws.row_dimensions[1].height = 30
        for i, port in enumerate(cfg['ports'], start=2):
            cell = ws.cell(row=i, column=1, value=port)
            cell.font      = Font(size=12)
            cell.alignment = Alignment(horizontal='left', vertical='center')
            ws.row_dimensions[i].height = 25
    wb.save(path)


def read_port_config(path):
    """Read port-config template. Returns {tab_name: [port, ...]} for any sheet found."""
    wb = openpyxl.load_workbook(path, data_only=True)
    result = {}
    for tab_name in PORT_TAB_CONFIG:
        if tab_name not in wb.sheetnames:
            continue
        ws = wb[tab_name]
        ports = []
        for row in ws.iter_rows(values_only=True):
            v = row[0] if row else None
            if v and str(v).strip() and str(v).strip() != '港口名称':
                ports.append(str(v).strip())
        if ports:
            result[tab_name] = ports
    return result


# ══════════════════════════════════════════════════════════════
# 样式常量
# ══════════════════════════════════════════════════════════════

_THIN = Side(style='thin')
BORDER_THIN = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
FILL_ORCHID  = PatternFill('solid', fgColor='DA70D6')
FILL_SUBTOTAL = PatternFill('solid', fgColor='FFF2CC')
ALIGN_C      = Alignment(horizontal='center', vertical='center', wrap_text=True)
ALIGN_L      = Alignment(horizontal='left',   vertical='center', wrap_text=True)
ALIGN_TOP_L  = Alignment(horizontal='left',   vertical='top')

SOURCE_TOP_H, SOURCE_HEADER_H, SOURCE_DATA_H = 17.45, 22.6, 23.25
SOURCE_SUBTOTAL_H, SOURCE_BLANK_H = 16.8, 17.45
SOURCE_FONT = Font(name='Arial', size=7)

def _f(size, bold=False):
    return Font(size=size, bold=bold)


def _apply(cell, font=None, alignment=None, fill=None, border=None):
    if font:      cell.font      = font
    if alignment: cell.alignment = alignment
    if fill:      cell.fill      = fill
    if border:    cell.border    = border


def _style_row(ws, row_num, n_cols, font, alignment, fill=None, border=None):
    for c in range(1, n_cols + 1):
        _apply(ws.cell(row=row_num, column=c), font=font, alignment=alignment,
               fill=fill, border=border)


def _style_source_row(ws, row_num, n_cols, height, border=True):
    ws.row_dimensions[row_num].height = height
    for col in range(1, n_cols + 1):
        _apply(
            ws.cell(row=row_num, column=col),
            font=SOURCE_FONT,
            alignment=ALIGN_C,
            border=BORDER_THIN if border else None,
        )


def _style_source_header(ws, row_num, n_cols, fixed_col_count):
    _style_source_row(ws, row_num, n_cols, SOURCE_HEADER_H)
    for col in range(1, min(fixed_col_count, n_cols) + 1):
        ws.cell(row=row_num, column=col).fill = FILL_ORCHID


def _style_source_subtotal(ws, row_num, n_cols, fixed_col_count, size_count):
    _style_source_row(ws, row_num, n_cols, SOURCE_SUBTOTAL_H)
    first_size_col = fixed_col_count + 1
    for col in range(first_size_col, first_size_col + size_count):
        ws.cell(row=row_num, column=col).fill = FILL_SUBTOTAL


def _style_source_blank(ws, row_num, n_cols):
    ws.row_dimensions[row_num].height = SOURCE_BLANK_H
    for col in range(1, n_cols + 1):
        _apply(ws.cell(row=row_num, column=col), font=SOURCE_FONT, alignment=ALIGN_TOP_L)


def _set_source_column_widths(ws, include_port, include_actual_dates,
                              actual_size_count=None):
    if actual_size_count is None:
        actual_size_count = SOURCE_SIZE_COLS
    first_size_col = 9 if include_port else 8
    total_col = first_size_col + actual_size_count
    date_col = total_col + 1
    note_col = total_col + (4 if include_actual_dates else 2)
    spacer_col = note_col + 1
    factory_col = note_col + 2

    if include_port:
        widths = {
            1: 8.86,
            7: 19.07,
            8: 10.53,
            first_size_col: 5.4,
            total_col: 7.73,
            date_col: 9.66,
            note_col: 22.93,
            spacer_col: 0.4,
            factory_col: 13.8,
        }
    else:
        widths = {
            1: 12.13,
            5: 19.86,
            6: 12.13,
            first_size_col: 6.13,
            total_col: 15.07,
            date_col: 10.8,
            note_col: 44.93,
            spacer_col: 0.4,
            factory_col: 13.4,
        }

    if include_actual_dates:
        widths[total_col + 2] = 10.8
        widths[total_col + 3] = 10.8

    for col, width in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = width


# ══════════════════════════════════════════════════════════════
# 配置读写
# ══════════════════════════════════════════════════════════════

def load_saved():
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding='utf-8'))
        except Exception:
            pass
    return {}


def save_saved(data):
    CONFIG_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')


def update_saved(**kv):
    """Merge key/values into the saved config without dropping existing keys."""
    data = load_saved()
    data.update(kv)
    save_saved(data)


# ══════════════════════════════════════════════════════════════
# 源工作表检测
# ══════════════════════════════════════════════════════════════

def detect_sources(wb):
    """Return (port_ws, qty_ws, suffix, ordered_pairs)."""
    port_ws = qty_ws = None
    suffix = ''
    ordered = []
    for name in wb.sheetnames:
        if '(2)' in name:
            continue
        if '港口明细汇总' in name and not port_ws:
            port_ws = wb[name]
            suffix = name.replace('港口明细汇总', '').strip()
            ordered.append((port_ws, name))
        elif '数量明细汇总' in name and not qty_ws:
            qty_ws = wb[name]
            ordered.append((qty_ws, name))
    if not port_ws or not qty_ws:
        raise ValueError(
            '未找到源工作表\n请确认文件包含"港口明细汇总"和"数量明细汇总"两张工作表'
        )
    return port_ws, qty_ws, suffix, ordered


# ══════════════════════════════════════════════════════════════
# 解析
# ══════════════════════════════════════════════════════════════

SOURCE_PORT_PREFIX = '港口明细汇总'
SOURCE_QTY_PREFIX = '数量明细汇总'
SOURCE_SIZE_COLS = 29

DESTINATION_PORT_BY_CODE = {
    '0101': '南非',
    '0104': '阿联酋',
    '0106': '英国伦敦',
    '1000': '美国',
    '1500': '巴西',
    '3000': '维多利亚',
    '7000': '新加坡',
    '8015': '中国深圳',
    '8094': '巴安斯',
    '8123': 'ASOS UK 英国',
    '8146': '新南威尔士',
    '8161': '英国诺丁汉',
    '8253': 'ASOS EU 欧洲',
    '9000': '新西兰',
}

# ══════════════════════════════════════════════════════════════
# 尺码排序
# ══════════════════════════════════════════════════════════════

_STD_SIZE_RANK = {
    'XXXXS': -6, '4XS': -6, 'XXXS': -5, '3XS': -5, 'XXS': -4, '2XS': -4, 'XS': -3,
    'S': 0, 'M': 2, 'L': 4, 'XL': 6, 'XXL': 7, '2XL': 7, 'XXXL': 8, '3XL': 8,
    'XXXXL': 9, '4XL': 9, '5XL': 10, '6XL': 11,
}


def _size_sort_key(s):
    u = str(s).upper().strip()
    if u in _STD_SIZE_RANK:
        return (0, _STD_SIZE_RANK[u], '')
    digits = ''.join(c for c in u if c.isdigit())
    letters = ''.join(c for c in u if not c.isdigit())
    if digits:
        return (1, int(digits), letters)
    return (2, 0, u)


DESTINATION_PORT_KEYWORDS = [
    ('ASOS EU', 'ASOS EU 欧洲'),
    ('ASOS UK', 'ASOS UK 英国'),
    ('ZALORA', '中国深圳'),
    ('BASH', '巴安斯'),
    ('ICONIC', '新南威尔士'),
    ('NEXT GROUP', '英国诺丁汉'),
    ('LIPSY', '英国诺丁汉'),
    ('SOUTH AFRICA', '南非'),
    ('UNITED ARAB EMIRATES', '阿联酋'),
    ('UAE', '阿联酋'),
    ('BRAZIL', '巴西'),
    ('BRASIL', '巴西'),
    ('UNITED STATES', '美国'),
    ('VICTORIAN', '维多利亚'),
    ('AUSTRALIA', '维多利亚'),
    ('SINGAPORE', '新加坡'),
    ('NEW ZEALAND', '新西兰'),
    ('UNITED KINGDOM', '英国伦敦'),
]


def _clean(v):
    if v is None:
        return ''
    return str(v).strip()


def _norm_filename(name):
    return ' '.join(name.replace('\u00a0', ' ').split())


def _po_from_filename(path):
    match = re.search(r'(?:^| - )(PO[^ -]+)(?= - )', _norm_filename(path.name))
    return match.group(1) if match else None


def latest_purchase_order_csvs(folder_path):
    """Return newest Purchase Order CSV per PO number, scanning subfolders too."""
    folder = Path(folder_path)
    if not folder.exists() or not folder.is_dir():
        raise ValueError(f'采购订单文件夹不存在：\n{folder}')

    by_po = {}
    for path in folder.rglob('*.csv'):
        if 'Purchase Order' not in _norm_filename(path.name):
            continue
        po_num = _po_from_filename(path)
        if not po_num:
            continue
        stat = path.stat()
        current = by_po.get(po_num)
        candidate = (stat.st_mtime, str(path), path)
        if current is None or candidate[:2] > current[:2]:
            by_po[po_num] = candidate

    if not by_po:
        raise ValueError(f'未在文件夹中找到 Purchase Order CSV：\n{folder}')
    return [item[2] for _, item in sorted(by_po.items())]


def _read_po_csv(path):
    for enc in ('utf-8-sig', 'utf-8', 'gbk'):
        try:
            with open(path, newline='', encoding=enc) as f:
                return list(csv.DictReader(f))
        except UnicodeDecodeError:
            continue
    with open(path, newline='') as f:
        return list(csv.DictReader(f))


def _split_style_colour(value):
    parts = _clean(value).split('-', 1)
    if len(parts) == 2:
        return parts[0].strip(), parts[1].strip()
    return _clean(value), ''


def _as_number(value):
    text = _clean(value).replace(',', '')
    if not text:
        return 0
    try:
        number = float(text)
    except ValueError:
        return 0
    return int(number) if number.is_integer() else number


def _date_to_iso(value):
    if isinstance(value, datetime):
        return value.strftime('%Y-%m-%d')
    if isinstance(value, date):
        return value.strftime('%Y-%m-%d')
    text = _clean(value)
    if not text:
        return ''
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y', '%m/%d/%Y'):
        try:
            return datetime.strptime(text, fmt).strftime('%Y-%m-%d')
        except ValueError:
            pass
    return text


def _port_from_destination(destination):
    text = _clean(destination)
    code = text.split(' - ', 1)[0].strip()
    if code in DESTINATION_PORT_BY_CODE:
        return DESTINATION_PORT_BY_CODE[code]
    upper = text.upper()
    for keyword, port in DESTINATION_PORT_KEYWORDS:
        if keyword in upper:
            return port
    return text or '未识别港口'


def _order_note(row, port=None, rec=None):
    marketing = _clean(row.get('MARKETING EVENT'))
    colour    = _clean(row.get('COLOUR DESCRIPTION'))
    dest_raw  = _clean(row.get('DESTINATION NAME', ''))

    parts = []

    # Wholesale note (requires port context)
    if port and marketing.upper() == 'WHOLESALE':
        # Prefer the full 港口代码 from 订单明细 (e.g. "BASH 8094");
        # fall back to the numeric code in DESTINATION NAME.
        code_disp = ''
        if rec:
            code_disp = str(rec.get('港口代码_detail') or '').strip()
        if not code_disp and ' - ' in dest_raw:
            code_disp = dest_raw.split(' - ', 1)[0].strip()
        note = '此合同是 WHOLESALE 订单'
        if code_disp:
            note += f'，{code_disp}'
        note += f'，{port}'
        parts.append(note)

    # Multi-colour note — use the Chinese colour name from 订单明细 when known
    colour_disp = colour
    if rec:
        rec_colour = str(rec.get('颜色') or '').strip()
        if rec_colour:
            colour_disp = rec_colour
    cn_parts = [p.strip() for p in colour_disp.split('/') if p.strip()]
    if len(cn_parts) > 1:
        n = len(cn_parts)
        style_code = ''
        if rec:
            style_code = str(rec.get('厂款号：') or '').strip()
        style_code = style_code or _clean(row.get('DEVELOPMENT CODE', ''))
        parts.append(
            f'此款是由{n}件 {style_code} 组成的{n}件套套装，配色是 {colour_disp}'
        )

    # 2-厂款号 set note (厂款号2 only known after field reference applied)
    if rec:
        fac2 = str(rec.get('厂款号2') or '').strip()
        if fac2:
            fac1 = str(rec.get('厂款号：') or '').strip()
            cust = str(rec.get('客款号') or '').strip()
            parts.append(
                f'客款号 {cust} 是由1件 {fac1} +1件 {fac2} 组成的套装。'
                f'洗水唛上的客款号和单件客款号不同。'
            )

    if not parts:
        return ''
    if len(parts) == 1:
        return parts[0]
    return '\n'.join(f'{i + 1}. {p}' for i, p in enumerate(parts))


FIELD_REFERENCE_COLUMNS = [
    '厂款号：', '厂款号2', '颜色', '内部色号', '英文颜色',
    '英文颜色拆分', '厂款号拆分', '工厂出厂期', '订单附加备注',
]
FIELD_REFERENCE_ALIASES = {
    '合同号：': ('合同号：', '合同号'),
    '厂款号：': ('厂款号：', '厂款号'),
}


def _rec_get(rec, field):
    for name in FIELD_REFERENCE_ALIASES.get(field, (field,)):
        value = rec.get(name)
        if value not in (None, ''):
            return value
    return None


def _reference_keys(po, customer_style, customer_colour, port=''):
    return [
        (po, customer_style, customer_colour, port),
        (po, customer_style, customer_colour, ''),
        ('', customer_style, customer_colour, port),
        ('', customer_style, customer_colour, ''),
    ]


def load_field_reference(path):
    """Read an existing source workbook and return product/color field lookups."""
    if not path:
        return {}
    ref_path = Path(path)
    if not ref_path.exists():
        return {}

    wb = openpyxl.load_workbook(ref_path, data_only=True)
    port_ws, qty_ws, _suffix, _ordered = detect_sources(wb)
    lookup = {}

    def add_record(rec, include_port):
        po = _clean(_rec_get(rec, '合同号：'))
        customer_style = _clean(_rec_get(rec, '客款号'))
        customer_colour = _clean(_rec_get(rec, '客户色号'))
        port = _clean(_rec_get(rec, '港口')) if include_port else ''
        if not customer_style or not customer_colour:
            return
        data = {}
        for field in FIELD_REFERENCE_COLUMNS:
            value = _rec_get(rec, field)
            if value not in (None, ''):
                data[field] = value
        for key in _reference_keys(po, customer_style, customer_colour, port):
            existing = lookup.setdefault(key, {})
            for field, value in data.items():
                existing.setdefault(field, value)

    for rec in parse_qty(qty_ws):
        add_record(rec, include_port=False)
    for rec in parse_port(port_ws):
        add_record(rec, include_port=True)

    return lookup


def _po_from_any_filename(path):
    match = re.search(r'PO\d+', _norm_filename(Path(path).name), re.IGNORECASE)
    return match.group(0).upper() if match else None


def _excel_files_under(path):
    root = Path(path)
    if not root.exists():
        return []
    if root.is_file():
        return [root] if root.suffix.lower() in ('.xlsx', '.xlsm', '.xls') else []
    return [
        p for p in root.rglob('*')
        if p.is_file()
        and p.suffix.lower() in ('.xlsx', '.xlsm', '.xls')
        and not p.name.startswith('~$')
    ]


def _detail_port_from_row(row):
    port = _clean(row.get('目的港'))
    if port:
        return port
    code_text = _clean(row.get('港口代码'))
    match = re.search(r'\b(\d{4})\b', code_text)
    if match and match.group(1) in DESTINATION_PORT_BY_CODE:
        return DESTINATION_PORT_BY_CODE[match.group(1)]
    return ''


def load_order_detail_reference(path):
    """Read 订单明细 files keyed by PO number in the filename."""
    if not path:
        return {}

    lookup = {}
    for file_path in _excel_files_under(path):
        po = _po_from_any_filename(file_path)
        if not po:
            continue
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
        except Exception:
            continue
        for ws in wb.worksheets:
            header_map = None
            for row in ws.iter_rows(values_only=True):
                values = [_clean(v) for v in row]
                if '客户色号' not in values:
                    continue
                header_map = {name: i for i, name in enumerate(values) if name}
                break
            if not header_map or '客户色号' not in header_map:
                continue

            for row in ws.iter_rows(min_row=2, values_only=True):
                row_values = list(row)
                detail_row = {
                    name: row_values[idx] if idx < len(row_values) else None
                    for name, idx in header_map.items()
                }
                customer_style, customer_colour = _split_style_colour(detail_row.get('客户色号'))
                if not customer_style or not customer_colour:
                    continue

                port = _detail_port_from_row(detail_row)
                data = {
                    '厂款号：': detail_row.get('厂款号'),
                    '厂款号2': detail_row.get('厂款号2'),
                    '工厂出厂期': _date_to_iso(detail_row.get('工厂出厂期')),
                    '颜色': detail_row.get('颜色名称'),
                    '内部色号': detail_row.get('颜色编号'),
                    '英文颜色': detail_row.get('英文颜色'),
                    '英文颜色拆分': detail_row.get('英文颜色拆分'),
                    '厂款号拆分': detail_row.get('厂款号拆分'),
                }
                data = {k: v for k, v in data.items() if v not in (None, '')}
                if not data:
                    continue

                detail_keys = [
                    (po, customer_style, customer_colour, port),
                    (po, customer_style, customer_colour, ''),
                ]
                for key in detail_keys:
                    existing = lookup.setdefault(key, {})
                    existing.update(data)
                # Tag the port-specific entry so _apply_field_reference can
                # override base['港口'] with the more accurate 目的港 value.
                if port:
                    port_entry = lookup[(po, customer_style, customer_colour, port)]
                    port_entry['目的港_detail'] = port
                    code_full = _clean(detail_row.get('港口代码'))
                    if code_full:
                        port_entry['港口代码_detail'] = code_full

    return lookup


def load_order_detail_components(path):
    """Extract set-pack component list from 订单明细 files.

    Returns {(po, customer_style, customer_colour): [component_dict, ...]}
    where len > 1 means a set-pack whose combined CSV row must be expanded.
    Each component_dict: {颜色, 内部色号, 英文颜色, 厂款号：, 工厂出厂期}
    Single-component items are included so their fields can also be applied.
    """
    if not path:
        return {}

    lookup = {}   # {(po, style, colour): [components]}

    for file_path in _excel_files_under(path):
        po = _po_from_any_filename(file_path)
        if not po:
            continue
        try:
            wb = openpyxl.load_workbook(file_path, data_only=True)
        except Exception:
            continue

        for ws in wb.worksheets:
            header_map = None
            for row in ws.iter_rows(values_only=True):
                vals = [_clean(v) for v in row]
                if '客户色号' not in vals:
                    continue
                header_map = {name: i for i, name in enumerate(vals) if name}
                break
            if not header_map or '客户色号' not in header_map:
                continue

            seen = {}   # {(po, style, colour): set of 颜色编号 already recorded}

            for row in ws.iter_rows(min_row=2, values_only=True):
                rv = list(row)
                d = {
                    name: rv[idx] if idx < len(rv) else None
                    for name, idx in header_map.items()
                }
                cstyle, ccolour = _split_style_colour(d.get('客户色号'))
                if not cstyle or not ccolour:
                    continue

                color_code = _clean(d.get('颜色编号', ''))
                key = (po, cstyle, ccolour)

                seen.setdefault(key, set())
                if color_code and color_code in seen[key]:
                    continue            # same color repeated for another port row
                if color_code:
                    seen[key].add(color_code)

                comp = {k: v for k, v in {
                    '颜色':      _clean(d.get('颜色名称', '')),
                    '内部色号':  color_code,
                    '英文颜色':  _clean(d.get('英文颜色', '')),
                    '英文颜色拆分': _clean(d.get('英文颜色拆分', '')),
                    '厂款号拆分': _clean(d.get('厂款号拆分', '')),
                    '厂款号：':  _clean(d.get('厂款号', '')),
                    '厂款号2':   _clean(d.get('厂款号2', '')),
                    '工厂出厂期': _date_to_iso(d.get('工厂出厂期', '')),
                }.items() if v}
                if not comp:
                    continue

                bucket = lookup.setdefault(key, [])
                existing_codes = {c.get('内部色号') for c in bucket}
                if color_code not in existing_codes:
                    bucket.append(comp)

    return lookup


def _merge_field_lookups(*lookups):
    merged = {}
    for lookup in lookups:
        for key, data in (lookup or {}).items():
            target = merged.setdefault(key, {})
            for field, value in data.items():
                if value not in (None, ''):
                    target[field] = value
    return merged


def _apply_field_reference(base, row, port, field_lookup):
    if not field_lookup:
        return base
    customer_style = base['客款号']
    customer_colour = base['客户色号']
    po = _clean(row.get('ORDER NUMBER'))
    for key in _reference_keys(po, customer_style, customer_colour, port or ''):
        ref = field_lookup.get(key)
        if not ref:
            continue
        for field in FIELD_REFERENCE_COLUMNS:
            value = ref.get(field)
            if value not in (None, ''):
                base[field] = value
        # When we matched via a port-specific key, use 订单明细's 目的港 as the
        # authoritative port (more accurate than the CSV DESTINATION_NAME mapping).
        if key[3]:
            detail_port = ref.get('目的港_detail')
            if detail_port:
                base['港口'] = detail_port
        code_full = ref.get('港口代码_detail')
        if code_full:
            base['港口代码_detail'] = code_full
        break
    return base


def _source_record_base(row, port=None, field_lookup=None):
    customer_style, customer_colour = _split_style_colour(row.get('STYLE COLOUR CODE'))
    description = _clean(row.get('STYLE DESCRIPTION'))
    # Use the same suffix rule as ①订单明细导入. A missing description gives no pack evidence.
    is_pack = None
    if description:
        is_pack = re.search(r' (?:SET|[0-9]*[1-9]PK)$', description, re.IGNORECASE) is not None
    base = {
        'is_pack': is_pack,
        '合同号：': _clean(row.get('ORDER NUMBER')),
        '客款号': customer_style,
        '客户色号': customer_colour,
        '厂款号：': _clean(row.get('DEVELOPMENT CODE')),
        '厂款号2': '',
        '颜色': _clean(row.get('COLOUR DESCRIPTION')),
        '内部色号': '',
        '英文颜色': _clean(row.get('COLOUR DESCRIPTION')),
        '港口': port or '',
        'destination_raw': _clean(row.get('DESTINATION NAME', '')),   # raw English destination name
        '工厂出厂期': _date_to_iso(row.get('FINAL BOOKED BY DATE')),
        '实际出厂期': _date_to_iso(row.get('FINAL HANDOVER DATE')),
        '实际船期': _date_to_iso(row.get('LATEST SHIP-BY DATE')),
        '订单附加备注': _order_note(row, port),
    }
    auto_note = base['订单附加备注']
    rec = _apply_field_reference(base, row, port, field_lookup)
    # Regenerate the note with field-reference knowledge (厂款号2, Chinese
    # colour, full 港口代码), unless an existing-workbook field reference
    # explicitly supplied its own note.
    if rec.get('订单附加备注') == auto_note:
        note_port = (rec.get('港口') or port) if port else None
        rec['订单附加备注'] = _order_note(row, note_port, rec)
    return rec


def _add_qty(target, size, qty):
    if not size:
        return
    target[size] += _as_number(qty)


_COLOUR_SEP_RE = re.compile(r'[，,/／、]')


def _split_colour_list(value):
    """把合并的颜色串按 中/英文逗号、斜杠、顿号 拆成单色列表。"""
    return [p.strip() for p in _COLOUR_SEP_RE.split(_clean(value)) if p.strip()]


def _en_pieces(rec):
    """逐件英文列表。优先用订单明细「英文颜色拆分」('|' 分隔，段内 / 保留、精确对齐)；
    该列为空时回退到按分隔符拆原始「英文颜色」。"""
    s = str(rec.get('英文颜色拆分') or '').strip()
    if s:
        return [p.strip() for p in s.split('|') if p.strip()]
    return _split_colour_list(rec.get('英文颜色', ''))


def _fac_pieces(rec):
    """逐件厂款号列表，来自订单明细「厂款号拆分」('|' 分隔，逐件与颜色/英文颜色对齐)；
    为空则返回 []（由调用方回退到旧的 厂款号/厂款号2 逻辑）。"""
    s = str(rec.get('厂款号拆分') or '').strip()
    return [p.strip() for p in s.split('|') if p.strip()] if s else []


def _expand_setpacks(records, component_lookup):
    """Expand set-pack records into one record per component.

    A set-pack is identified when (po, customer_style, customer_colour) maps
    to N > 1 entries in component_lookup (e.g. a 3-colour lingerie set).
    Each CSV combined row becomes N rows with the same sizes replicated —
    one of each colour ships per pack, so quantities per component equal the
    pack quantity.
    Single-component entries are also used to apply the more accurate
    颜色/内部色号/厂款号 fields from 订单明细.
    """
    if not component_lookup:
        return records

    result = []
    for rec in records:
        po      = _clean(rec.get('合同号：', ''))
        cstyle  = _clean(rec.get('客款号', ''))
        ccolour = _clean(rec.get('客户色号', ''))
        comps   = component_lookup.get((po, cstyle, ccolour), [])

        if not comps:
            result.append(rec)
            continue

        # Apply component fields to every expanded record (or the single record)
        n = len(comps)
        # 套装(N>1)时，英文颜色常以合并串形式（如 "MOON TIDE, PINK KISS,
        # CHEEKY CHERRY"）出现在每个组件上，需按分量拆分对齐到第 i 个组件 ——
        # 与中文颜色（颜色名称）逐件拆分保持一致。
        base_en_parts = _en_pieces(rec)
        for i, comp in enumerate(comps):
            new_rec = dict(rec)
            if n > 1:
                new_rec['is_pack'] = True
            for field in ('颜色', '内部色号', '工厂出厂期'):
                if comp.get(field):
                    new_rec[field] = comp[field]
            # 英文颜色：组件自带优先按分量拆，其次按基础合并串拆
            comp_en_parts = _en_pieces(comp)
            if n > 1 and len(comp_en_parts) == n:
                new_rec['英文颜色'] = comp_en_parts[i]
            elif n > 1 and len(base_en_parts) == n:
                new_rec['英文颜色'] = base_en_parts[i]
            elif comp.get('英文颜色'):
                new_rec['英文颜色'] = comp['英文颜色']
            if comp.get('厂款号：'):
                new_rec['厂款号：'] = comp['厂款号：']
            if comp.get('厂款号2'):
                new_rec['厂款号2'] = comp['厂款号2']
            result.append(new_rec)

    return result


def _records_from_po_csvs(csv_paths, field_lookup=None, component_lookup=None):
    port_groups = OrderedDict()
    qty_groups = OrderedDict()

    for path in csv_paths:
        rows = _read_po_csv(path)
        for row in rows:
            po = _clean(row.get('ORDER NUMBER'))
            if not po:
                continue
            customer_style, customer_colour = _split_style_colour(row.get('STYLE COLOUR CODE'))
            common_key = (
                po,
                customer_style,
                customer_colour,
                _clean(row.get('DEVELOPMENT CODE')),
                _clean(row.get('COLOUR DESCRIPTION')),
            )
            port = _port_from_destination(row.get('DESTINATION NAME'))

            if common_key not in qty_groups:
                qty_groups[common_key] = {
                    'base': _source_record_base(row, field_lookup=field_lookup),
                    'sizes': defaultdict(int),
                }
            _add_qty(qty_groups[common_key]['sizes'], _clean(row.get('SIZE')), row.get('QUANTITY'))

            port_key = common_key + (port,)
            if port_key not in port_groups:
                port_groups[port_key] = {
                    'base': _source_record_base(row, port, field_lookup=field_lookup),
                    'sizes': defaultdict(int),
                }
            _add_qty(port_groups[port_key]['sizes'], _clean(row.get('SIZE')), row.get('QUANTITY'))

    qty_records = []
    for item in qty_groups.values():
        rec = dict(item['base'])
        rec['sizes_by_name'] = dict(item['sizes'])
        qty_records.append(rec)

    port_records = []
    for item in port_groups.values():
        rec = dict(item['base'])
        rec['sizes_by_name'] = dict(item['sizes'])
        port_records.append(rec)

    # Expand set-pack combined rows into one record per colour component.
    port_records = _expand_setpacks(port_records, component_lookup)
    qty_records  = _expand_setpacks(qty_records,  component_lookup)

    return port_records, qty_records


def _ordered_sizes(records):
    sizes = []
    for rec in records:
        for size in rec['sizes_by_name']:
            if size not in sizes:
                sizes.append(size)
    sizes.sort(key=_size_sort_key)
    return sizes[:SOURCE_SIZE_COLS]


def _expand_rec_fac(rec, keep_both):
    """Use 厂款号2 when set, and include the first style when the record needs expanding."""
    fac2 = str(rec.get('厂款号2') or '').strip()
    if not fac2:
        return [rec]
    second = {**rec, '厂款号：': fac2}
    return [rec, second] if keep_both else [second]


def _expand_rec_colours(rec):
    """If 颜色 contains '/', return one copy per part with 颜色/英文颜色 split."""
    colour    = str(rec.get('颜色') or '').strip()
    en_colour = str(rec.get('英文颜色') or '').strip()
    cn_parts  = [p.strip() for p in colour.split('/') if p.strip()]
    en_parts  = _en_pieces(rec)                         # '|' 拆分优先，回退分隔符
    if len(cn_parts) <= 1:
        return [rec]
    result = []
    for i, cn_p in enumerate(cn_parts):
        en_p = en_parts[i] if i < len(en_parts) else en_colour
        new_rec = dict(rec)
        new_rec['颜色'] = cn_p
        new_rec['英文颜色'] = en_p
        result.append(new_rec)
    return result


def _expand_rec_fac_colours(rec):
    """展开一条记录为逐件记录，厂款号与颜色/英文颜色按索引一一对齐。

    若「厂款号拆分」存在（步骤①拆分弹窗逐件填写，'|' 分隔）且颜色分段 >1，
    直接按索引配对——每件各自的颜色/英文颜色/厂款号，支持任意 N 件、任意厂款号组合
    （如 3 件中 2 件同厂款号、1 件不同）。套装或多色仍按厂款号和颜色展开。
    CSV 明确为普通单色款时，已有厂款号2就只输出该厂款号，数量只计一次。
    """
    fac_parts = _fac_pieces(rec)
    colour = str(rec.get('颜色') or '').strip()
    cn_parts = [p.strip() for p in colour.split('/') if p.strip()]
    if fac_parts and len(cn_parts) > 1:
        en_parts = _en_pieces(rec)
        en_colour = str(rec.get('英文颜色') or '').strip()
        result = []
        for i, cn_p in enumerate(cn_parts):
            new_rec = dict(rec)
            new_rec['颜色'] = cn_p
            new_rec['英文颜色'] = en_parts[i] if i < len(en_parts) else en_colour
            new_rec['厂款号：'] = fac_parts[i] if i < len(fac_parts) else fac_parts[-1]
            result.append(new_rec)
        return result
    keep_both = len(cn_parts) > 1 or rec.get('is_pack') is not False
    return [c for f in _expand_rec_fac(rec, keep_both) for c in _expand_rec_colours(f)]


def _write_source_sheet(ws, records, include_port, include_actual_dates=False):
    fixed_headers = ['合同号：', '客款号', '客户色号', '厂款号：', '颜色', '内部色号', '英文颜色']
    if include_port:
        fixed_headers.append('港口')
    tail_headers = ['总数', '工厂出厂期']
    if include_actual_dates:
        tail_headers += ['实际出厂期', '实际船期']
    tail_headers += ['订单附加备注']

    def make_header(size_headers):
        local_sizes = size_headers[:SOURCE_SIZE_COLS]
        header = fixed_headers + local_sizes + tail_headers
        return header, local_sizes

    groups = OrderedDict()
    for rec in records:
        groups.setdefault(rec['合同号：'], []).append(rec)

    # Pre-compute size headers per group to find max width for top blank row
    group_data = []
    max_actual_size_count = 0
    for group in groups.values():
        gsz = _add_fixed_endpoints(_ordered_sizes(group))
        if len(gsz) > max_actual_size_count:
            max_actual_size_count = len(gsz)
        group_data.append((group, gsz))

    top_row_width = len(fixed_headers) + max_actual_size_count + len(tail_headers)
    ws.append([None] * (top_row_width + 1) + ['工厂'])
    top_row = ws.max_row
    _style_source_blank(ws, top_row, top_row_width + 2)

    for group, group_size_headers in group_data:
        header, size_headers = make_header(group_size_headers)
        ws.append(header)
        hdr_row = ws.max_row
        _style_source_header(ws, hdr_row, len(header), fixed_col_count=7)

        color_groups = OrderedDict()
        if include_port:
            # In the port summary, each PO/color can span multiple destination
            # ports. Keep those rows together and add one subtotal row per color.
            key_headers = fixed_headers[:-1]
            for rec in group:
                color_key = tuple(rec.get(h, '') for h in key_headers)
                color_groups.setdefault(color_key, []).append(rec)
            output_groups = color_groups.values()
        else:
            output_groups = [group]

        for output_group in output_groups:
            data_start = ws.max_row + 1
            for raw_rec in output_group:
                for rec in _expand_rec_fac_colours(raw_rec):
                    values = [rec.get(h, '') for h in fixed_headers]
                    size_values = [rec['sizes_by_name'].get(size, '') for size in size_headers]
                    total = sum(v for v in size_values if isinstance(v, (int, float)))
                    values += size_values + [total, rec.get('工厂出厂期', '')]
                    if include_actual_dates:
                        values += [rec.get('实际出厂期', ''), rec.get('实际船期', '')]
                    values += [rec.get('订单附加备注', '')]
                    ws.append(values)
                    row_num = ws.max_row
                    _style_source_row(ws, row_num, len(header), SOURCE_DATA_H)

            if include_port:
                data_end = ws.max_row
                subtotal_row = ws.max_row + 1
                first_size_col = len(fixed_headers) + 1
                subtotal_sizes = []
                for i, _size in enumerate(size_headers):
                    col = get_column_letter(first_size_col + i)
                    subtotal_sizes.append(f'=SUM({col}{data_start}:{col}{data_end})')
                if size_headers:
                    first_col = get_column_letter(first_size_col)
                    last_col = get_column_letter(first_size_col + len(size_headers) - 1)
                    subtotal_total = f'=SUM({first_col}{subtotal_row}:{last_col}{subtotal_row})'
                else:
                    subtotal_total = ''
                subtotal = (
                    [None] * len(fixed_headers)
                    + subtotal_sizes
                    + [subtotal_total]
                    + [None] * (len(tail_headers) - 1)
                )
                ws.append(subtotal)
                row_num = ws.max_row
                _style_source_subtotal(
                    ws,
                    row_num,
                    len(header),
                    fixed_col_count=len(fixed_headers),
                    size_count=len(size_headers),
                )

        ws.append([None] * len(header))
        _style_source_blank(ws, ws.max_row, len(header))

    _set_source_column_widths(ws, include_port, include_actual_dates, max_actual_size_count)


def build_source_workbook_from_po_folder(
    folder_path, year_quarter, field_source_path=None, include_actual_dates=False,
    order_detail_path=None,
):
    csv_paths = latest_purchase_order_csvs(folder_path)
    field_lookup = _merge_field_lookups(
        load_field_reference(field_source_path),
        load_order_detail_reference(order_detail_path),
    )
    component_lookup = load_order_detail_components(order_detail_path)
    port_records, qty_records = _records_from_po_csvs(
        csv_paths,
        field_lookup=field_lookup,
        component_lookup=component_lookup,
    )
    if not port_records or not qty_records:
        raise ValueError('采购订单 CSV 中没有可生成的明细数据')

    suffix = _clean(year_quarter)
    wb = Workbook()
    wb.remove(wb.active)
    port_ws = wb.create_sheet(f'{SOURCE_PORT_PREFIX}{suffix}')
    qty_ws = wb.create_sheet(f'{SOURCE_QTY_PREFIX}{suffix}')
    _write_source_sheet(
        port_ws, port_records,
        include_port=True,
        include_actual_dates=include_actual_dates,
    )
    _write_source_sheet(
        qty_ws, qty_records,
        include_port=False,
        include_actual_dates=include_actual_dates,
    )
    return wb


def source_summary_output_path(output_path, year_quarter):
    today = datetime.now()
    date_tag = f'{today.year}.{today.month}.{today.day}'
    yq = _clean(year_quarter) or '源数据'
    filename = f'CO {yq}  港口明细+数量明细汇总 {date_tag}.xlsx'
    return Path(output_path).parent / filename


def _scan_sizes(header_row, start, stop):
    names, indices = [], []
    for i in range(start, min(stop, len(header_row))):
        v = header_row[i]
        if v is not None and str(v).strip():
            names.append(str(v).strip())
            indices.append(i)
        else:
            break
    return names, indices


def _find_col(header_row, names, default=None):
    wanted = set(names)
    for i, value in enumerate(header_row):
        if _clean(value) in wanted:
            return i
    return default


def _row_get(row, idx):
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def parse_port(ws):
    TOTAL, DATE, NOTE = 37, 38, 39
    ACTUAL_DATE = ACTUAL_SHIP = None
    sh, si = [], list(range(8, 15))
    records = []
    for row in ws.iter_rows(values_only=True):
        if not any(v is not None for v in row):
            continue
        if row[0] == '合同号：':
            TOTAL = _find_col(row, ['总数'], TOTAL)
            DATE = _find_col(row, ['工厂出厂期'], DATE)
            ACTUAL_DATE = _find_col(row, ['实际出厂期'], ACTUAL_DATE)
            ACTUAL_SHIP = _find_col(row, ['实际船期'], ACTUAL_SHIP)
            NOTE = _find_col(row, ['订单附加备注'], NOTE)
            sh, si = _scan_sizes(row, 8, TOTAL)
            continue
        if not row[0] or str(row[0]).strip() == '':
            continue
        records.append({
            '合同号':       row[0],
            '客款号':       row[1],
            '客户色号':     row[2],
            '厂款号':       row[3],
            '颜色':         row[4],
            '内部色号':     row[5],
            '英文颜色':     row[6],
            '港口':         row[7],
            'size_headers': sh[:],
            'sizes':        [row[i] for i in si],
            '总数':         _row_get(row, TOTAL),
            '工厂出厂期':   _row_get(row, DATE),
            '实际出厂期':   _row_get(row, ACTUAL_DATE),
            '实际船期':     _row_get(row, ACTUAL_SHIP),
            '订单附加备注': _row_get(row, NOTE),
        })
    return records


def parse_qty(ws):
    TOTAL, DATE, NOTE = 36, 37, 38
    ACTUAL_DATE = ACTUAL_SHIP = None
    sh, si = [], list(range(7, 14))
    records = []
    for row in ws.iter_rows(values_only=True):
        if not any(v is not None for v in row):
            continue
        if row[0] == '合同号：':
            TOTAL = _find_col(row, ['总数'], TOTAL)
            DATE = _find_col(row, ['工厂出厂期'], DATE)
            ACTUAL_DATE = _find_col(row, ['实际出厂期'], ACTUAL_DATE)
            ACTUAL_SHIP = _find_col(row, ['实际船期'], ACTUAL_SHIP)
            NOTE = _find_col(row, ['订单附加备注'], NOTE)
            sh, si = _scan_sizes(row, 7, TOTAL)
            continue
        if not row[0] or str(row[0]).strip() == '':
            continue
        records.append({
            '合同号':       row[0],
            '客款号':       row[1],
            '客户色号':     row[2],
            '厂款号':       row[3],
            '颜色':         row[4],
            '内部色号':     row[5],
            '英文颜色':     row[6],
            'size_headers': sh[:],
            'sizes':        [row[i] for i in si],
            '总数':         _row_get(row, TOTAL),
            '工厂出厂期':   _row_get(row, DATE),
            '实际出厂期':   _row_get(row, ACTUAL_DATE),
            '实际船期':     _row_get(row, ACTUAL_SHIP),
            '订单附加备注': _row_get(row, NOTE),
        })
    return records


# ══════════════════════════════════════════════════════════════
# 工作表生成
# ══════════════════════════════════════════════════════════════

def copy_tab(src_ws, dest_wb, tab_name):
    ws = dest_wb.create_sheet(title=tab_name)
    for ltr, dim in src_ws.column_dimensions.items():
        ws.column_dimensions[ltr].width = dim.width
    for idx, dim in src_ws.row_dimensions.items():
        ws.row_dimensions[idx].height = dim.height
    for row in src_ws.iter_rows():
        for cell in row:
            nc = ws.cell(row=cell.row, column=cell.column, value=cell.value)
            if cell.has_style:
                nc.font          = copy(cell.font)
                nc.border        = copy(cell.border)
                nc.fill          = copy(cell.fill)
                nc.number_format = cell.number_format
                nc.protection    = copy(cell.protection)
                nc.alignment     = copy(cell.alignment)
    for merge in src_ws.merged_cells.ranges:
        ws.merge_cells(str(merge))


def gen_qty_by_color(dest_wb, qty_records):
    """数量清单按颜色: one block per (颜色, 厂款号) — header (orchid A-G) + data + subtotal.

    Sizes vary by garment style (厂款号): 6 / 7 / 8 / 11 size columns all occur.
    A color (e.g. 沁粉色) can contain several styles with DIFFERENT size schemes,
    so each style gets its own block and its own size header — otherwise the data
    columns and the 合计 subtotal would be misaligned. Subtotal label is still
    `{颜色}合计` (a color may repeat across styles, matching the sample file).
    """
    ws = dest_wb.create_sheet(title='数量清单按颜色')

    # Determine max sizes for column-width setup
    max_n = max((len(r['size_headers']) for r in qty_records), default=7)
    last_size_col  = 7 + max_n          # 1-indexed Excel col of last size
    total_col      = 8 + max_n          # 总数
    notes_col      = 10 + max_n         # 订单附加备注

    # Column widths
    ws.column_dimensions['A'].width = 18.73
    ws.column_dimensions[get_column_letter(3)].width  = 10.73   # 客户色号
    ws.column_dimensions[get_column_letter(4)].width  = 18.73   # 厂款号
    ws.column_dimensions[get_column_letter(8)].width  = 10.73   # first size
    ws.column_dimensions[get_column_letter(last_size_col)].width = 10.4
    ws.column_dimensions[get_column_letter(total_col)].width     = 10.73
    ws.column_dimensions[get_column_letter(notes_col)].width     = 113.6

    # Group by (颜色, 厂款号). Order: color by first appearance, then 厂款号 by
    # first appearance within that color (stable sort keeps same-color blocks
    # together while preserving source order of styles).
    color_order, groups, key_order = [], {}, []
    for rec in qty_records:
        key = (rec['颜色'], rec['厂款号'])
        if rec['颜色'] not in color_order:
            color_order.append(rec['颜色'])
        if key not in groups:
            groups[key] = []
            key_order.append(key)
        groups[key].append(rec)
    key_order.sort(key=lambda k: color_order.index(k[0]))

    for color, style in key_order:
        grp = groups[(color, style)]

        # Unified size headers: sorted union of every record's size_headers.
        # Records from different POs can have different size schemes (e.g. some
        # have 2XS/2XL, others don't). Using only grp[0] would misalign values.
        _all_sh: list = []
        for _r in grp:
            for _s in _r['size_headers']:
                if _s not in _all_sh:
                    _all_sh.append(_s)
        _all_sh.sort(key=_size_sort_key)
        sh = _add_fixed_endpoints(_all_sh)   # add 2XS/2XL and 10A/14D endpoints
        n   = len(sh)
        n_cols = 10 + n  # total cols this group outputs

        # ── Header row ──────────────────────────────────────
        ws.append(
            ['合同号：', '客款号', '客户色号', '厂款号：', '颜色', '内部色号', '英文颜色']
            + sh + ['总数', '工厂出厂期', '订单附加备注']
        )
        hdr_row = ws.max_row
        ws.row_dimensions[hdr_row].height = 22
        # orchid fill on cols A-G, plain style on size/totals cols
        for c in range(1, 8):
            _apply(ws.cell(row=hdr_row, column=c),
                   font=_f(11), alignment=ALIGN_C, fill=FILL_ORCHID, border=BORDER_THIN)
        for c in range(8, n_cols + 1):
            _apply(ws.cell(row=hdr_row, column=c),
                   font=_f(11), alignment=ALIGN_C, border=BORDER_THIN)

        data_start = ws.max_row + 1

        # ── Data rows ────────────────────────────────────────
        for rec in grp:
            # Look up each group-level size by name from this record's own
            # size_headers/sizes pair so mismatched schemes align correctly.
            _rec_sz = dict(zip(rec['size_headers'], rec['sizes']))
            _size_vals = [_rec_sz.get(s, '') for s in sh]
            ws.append(
                [rec['合同号'], rec['客款号'], rec['客户色号'], rec['厂款号'],
                 rec['颜色'], rec['内部色号'], rec['英文颜色']]
                + _size_vals
                + [rec['总数'], rec['工厂出厂期'], rec['订单附加备注']]
            )
            r = ws.max_row
            ws.row_dimensions[r].height = 18
            _style_row(ws, r, n_cols, font=_f(11), alignment=ALIGN_C, border=BORDER_THIN)

        data_end = ws.max_row

        # ── Subtotal row ─────────────────────────────────────
        sub = [f'{color}合计'] + [None] * 6
        for i in range(n):
            cl = get_column_letter(8 + i)
            sub.append(f'=SUM({cl}{data_start}:{cl}{data_end})')
        tcl = get_column_letter(8 + n)
        sub.append(f'=SUM({tcl}{data_start}:{tcl}{data_end})')
        ws.append(sub)

        sub_row = ws.max_row
        ws.row_dimensions[sub_row].height = 22
        _style_row(ws, sub_row, n_cols, font=_f(11), alignment=ALIGN_C, border=BORDER_THIN)
        # Bold label, merge A:G
        ws.cell(row=sub_row, column=1).font = _f(11, bold=True)
        ws.merge_cells(start_row=sub_row, start_column=1,
                       end_row=sub_row, end_column=7)

        # ── Blank spacer between sections ────────────────────
        ws.append([None] * n_cols)
        ws.row_dimensions[ws.max_row].height = 8


# ══════════════════════════════════════════════════════════════
# 颜色套装港口汇总（按 颜色 × 套装性 × 港口类别 各一条合计行）
# ══════════════════════════════════════════════════════════════

# 特殊港口集合（英国诺丁汉 + 特殊港口页的三个港口）
_SPECIAL_PORTS = {'英国诺丁汉', '英国伦敦', 'ASOS UK 英国', 'ASOS EU 欧洲'}

# 港口类别 → 备注末尾标记
_PORT_MARKER = {'special': '带※号', 'brazil': '尺码牌不一样', 'normal': ''}

# 件数：从备注解析「N件套/N件装」（中文或数字）
_PIECE_RE = re.compile(r'([0-9]+|[一二三四五六七八九十两]+)\s*件[套装]')


def _port_class(port):
    """港口三分类：巴西 / 特殊(诺丁汉+特殊港口页) / 普通(其余)。"""
    p = str(port or '').strip()
    return 'brazil' if p == '巴西' else ('special' if p in _SPECIAL_PORTS else 'normal')


def _piece_label(note):
    """件数标签：备注含「N件套/件装」→「N件装」；否则「单件装」。"""
    m = _PIECE_RE.search(str(note or ''))
    if m and m.group(1) not in ('1', '一', '单'):
        return f'{m.group(1)}件装'
    return '单件装'


def gen_colour_class_summary(dest_wb, port_records):
    """颜色套装港口汇总（明细式）：按 颜色 分节。

    列（前 7 列表头填粉）：合同号 客款号 客户色号 厂款号 颜色 内部色号 英文颜色 | 港口 |
    尺码… | 总数 | 工厂出厂期 | 备注。同一「行 + 同一港口类别」的多个港口用 / 合并、尺码求和；
    备注 = 「{件数}件装 客款号是{客款号}[ 港口标记]」（件数取自备注 N件套，港口标记按类别）。
    每颜色节末一条 {颜色}合计 SUM 行。
    """
    ws = dest_wb.create_sheet(title='颜色套装港口汇总')

    # 按 (颜色, 厂款号) 分节：颜色按首现，同色内 厂款号 按首现（稳定排序聚拢同色）——
    # 与「数量清单按颜色」一致，不同厂款号各自成节（尺码方案常不同）。
    color_order, groups, key_order = [], {}, []
    for rec in port_records:
        key = (rec['颜色'], rec['厂款号'])
        if rec['颜色'] not in color_order:
            color_order.append(rec['颜色'])
        if key not in groups:
            groups[key] = []
            key_order.append(key)
        groups[key].append(rec)
    key_order.sort(key=lambda k: color_order.index(k[0]))

    # 各节尺码并集（预计算，供列宽 max_n）
    block_sh = {}
    for key in key_order:
        raw = []
        for r in groups[key]:
            for s in r['size_headers']:
                if s not in raw:
                    raw.append(s)
        raw.sort(key=_size_sort_key)
        block_sh[key] = _add_fixed_endpoints(raw)
    max_n = max((len(sh) for sh in block_sh.values()), default=7)

    # 列宽
    for ci, w in ((1, 12), (2, 10), (3, 9), (4, 12), (5, 10), (6, 13), (7, 18), (8, 14)):
        ws.column_dimensions[get_column_letter(ci)].width = w
    for ci in range(9, 9 + max_n):
        ws.column_dimensions[get_column_letter(ci)].width = 6
    ws.column_dimensions[get_column_letter(9 + max_n)].width = 9      # 总数
    ws.column_dimensions[get_column_letter(10 + max_n)].width = 13    # 工厂出厂期
    ws.column_dimensions[get_column_letter(11 + max_n)].width = 32    # 备注

    HEAD = ['合同号', '客款号', '客户色号', '厂款号', '颜色', '内部色号', '英文颜色', '港口']

    for colour, factory in key_order:
        recs = groups[(colour, factory)]
        sh = block_sh[(colour, factory)]
        n = len(sh)
        n_cols = 11 + n

        # ── 列头行（1-7 列填粉）────────────────────────────
        ws.append(HEAD + sh + ['总数', '工厂出厂期', '备注'])
        r_hdr = ws.max_row
        ws.row_dimensions[r_hdr].height = 20
        _style_row(ws, r_hdr, n_cols, font=_f(11, bold=True), alignment=ALIGN_C,
                   border=BORDER_THIN)
        for c in range(1, 8):
            ws.cell(r_hdr, c).fill = FILL_ORCHID

        # ── 合并分组：同款/色号/内部色号/英文/出厂期 + 同港口类别（不分 合同号）──
        #    合同号 与 港口 均 / 合并、尺码/总数求和。
        order, merged = [], {}
        for r in recs:
            key = (r['客款号'], r['客户色号'], r['厂款号'],
                   r['内部色号'], r['英文颜色'], str(r['工厂出厂期'] or ''),
                   _port_class(r['港口']))
            if key not in merged:
                merged[key] = []
                order.append(key)
            merged[key].append(r)

        data_start = ws.max_row + 1
        for key in order:
            grp = merged[key]
            pclass = key[6]
            pos, ports = [], []
            for r in grp:
                po = str(r['合同号'] or '').strip()
                if po and po not in pos:
                    pos.append(po)
                p = str(r['港口'] or '').strip()
                if p and p not in ports:
                    ports.append(p)
            acc, total = {}, 0
            for r in grp:
                for s, v in zip(r['size_headers'], r['sizes']):
                    if isinstance(v, (int, float)):
                        acc[s] = acc.get(s, 0) + v
                t = r['总数']
                if isinstance(t, (int, float)):
                    total += t
            size_vals = [acc.get(s) or None for s in sh]
            g0 = grp[0]
            note_all = ' '.join(str(r.get('订单附加备注') or '') for r in grp)
            marker = _PORT_MARKER.get(pclass, '')
            remark = '%s 客款号是%s%s' % (
                _piece_label(note_all), g0['客款号'],
                (' ' + marker) if marker else '')
            ws.append(
                ['/'.join(pos), g0['客款号'], g0['客户色号'], g0['厂款号'],
                 colour, g0['内部色号'], g0['英文颜色'], '/'.join(ports)]
                + size_vals
                + [total or None, g0['工厂出厂期'], remark])
            r_data = ws.max_row
            ws.row_dimensions[r_data].height = 20
            _style_row(ws, r_data, n_cols, font=_f(11), alignment=ALIGN_C,
                       border=BORDER_THIN)
            ws.cell(r_data, n_cols).alignment = ALIGN_L      # 备注左对齐
        data_end = ws.max_row

        # ── {颜色}合计（SUM 该颜色全部明细行）──────────────
        if data_end >= data_start:
            total_row = ['%s合计' % colour] + [None] * 7
            for i in range(n):
                cl = get_column_letter(9 + i)
                total_row.append(f'=SUM({cl}{data_start}:{cl}{data_end})')
            tcl = get_column_letter(9 + n)
            total_row.append(f'=SUM({tcl}{data_start}:{tcl}{data_end})')
            total_row += [None, None]                        # 工厂出厂期 / 备注
            ws.append(total_row)
            r_tot = ws.max_row
            ws.row_dimensions[r_tot].height = 20
            _style_row(ws, r_tot, n_cols, font=_f(11, bold=True), alignment=ALIGN_C,
                       fill=FILL_SUBTOTAL2, border=BORDER_THIN)
            ws.merge_cells(start_row=r_tot, start_column=1, end_row=r_tot, end_column=8)

        # ── 空行分隔 ────────────────────────────────────────
        ws.append([None] * n_cols)
        ws.row_dimensions[ws.max_row].height = 8


# ══════════════════════════════════════════════════════════════
# 损耗汇总（外贸跟单第三张表）
# ══════════════════════════════════════════════════════════════

def _calc_loss(qty):
    """Wastage allowance based on order quantity tier.

    Tiers (per row total):
        1–10    → +3
        11–30   → +6
        31–100  → +8
        101–300 → +10
        301–500 → +15
        501–800 → +20
        801–1000→ +25
        1000+   → ceil(qty × 3 %)
    """
    try:
        q = float(qty or 0)
    except (TypeError, ValueError):
        return 0
    if q <= 0:   return 0
    if q <= 10:  return 3
    if q <= 30:  return 6
    if q <= 100: return 8
    if q <= 300: return 10
    if q <= 500: return 15
    if q <= 800: return 20
    if q <= 1000: return 25
    return math.ceil(q * 0.03)


def gen_loss_by_style(dest_wb, qty_records_a):
    """Add 损耗汇总 sheet to dest_wb.

    Groups format-A qty records (sizes_by_name dict) by 厂款号, shows one row
    per (PO, colour) record, and appends 损耗 + 含损耗合计 columns whose values
    are calculated from each row's 总数 using _calc_loss().

    A global size-column layout (sorted union of all sizes across every record)
    ensures 总数 / 损耗 / 含损耗合计 always land in the same columns regardless
    of which style a record belongs to.
    """
    ws = dest_wb.create_sheet(title='损耗汇总')

    # ── Global sorted size union ──────────────────────────────
    global_sh: list = []
    for rec in qty_records_a:
        for s in rec.get('sizes_by_name', {}):
            if s not in global_sh:
                global_sh.append(s)
    global_sh.sort(key=_size_sort_key)
    gn = len(global_sh)

    # Fixed column positions (1-indexed)
    total_col = 8 + gn        # 总数
    loss_col  = 8 + gn + 1   # 损耗
    incl_col  = 8 + gn + 2   # 含损耗合计
    n_cols    = 8 + gn + 2   # last column index (1-indexed)

    # ── Column widths ────────────────────────────────────────
    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 10
    ws.column_dimensions['C'].width = 10
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 16
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 24
    ws.column_dimensions[get_column_letter(total_col)].width = 8
    ws.column_dimensions[get_column_letter(loss_col)].width  = 7
    ws.column_dimensions[get_column_letter(incl_col)].width  = 11

    # ── Group by 厂款号 ───────────────────────────────────────
    style_order: list = []
    groups: dict = {}
    for rec in qty_records_a:
        key = rec.get('厂款号：', '') or rec.get('厂款号', '')
        if key not in groups:
            groups[key] = []
            style_order.append(key)
        groups[key].append(rec)

    for style in style_order:
        grp = groups[style]

        # ── Header ───────────────────────────────────────────
        ws.append(
            ['合同号：', '客款号', '客户色号', '厂款号：', '颜色', '内部色号', '英文颜色']
            + global_sh + ['总数', '损耗', '含损耗合计']
        )
        hdr_row = ws.max_row
        ws.row_dimensions[hdr_row].height = 22
        for c in range(1, 8):
            _apply(ws.cell(row=hdr_row, column=c),
                   font=_f(11), alignment=ALIGN_C, fill=FILL_ORCHID, border=BORDER_THIN)
        for c in range(8, n_cols + 1):
            _apply(ws.cell(row=hdr_row, column=c),
                   font=_f(11), alignment=ALIGN_C, border=BORDER_THIN)

        data_start = ws.max_row + 1

        # ── Data rows ─────────────────────────────────────────
        for rec in grp:
            szm  = rec.get('sizes_by_name', {})
            svs  = [szm.get(s, '') for s in global_sh]
            total = sum(v for v in szm.values() if isinstance(v, (int, float)))
            loss  = _calc_loss(total)
            ws.append(
                [rec.get('合同号：', ''), rec.get('客款号', ''), rec.get('客户色号', ''),
                 rec.get('厂款号：', ''), rec.get('颜色', ''), rec.get('内部色号', ''),
                 rec.get('英文颜色', '')]
                + svs
                + [total, loss, total + loss]
            )
            r = ws.max_row
            ws.row_dimensions[r].height = 18
            _style_row(ws, r, n_cols, font=_f(11), alignment=ALIGN_C, border=BORDER_THIN)

        data_end = ws.max_row

        # ── Subtotal ──────────────────────────────────────────
        sub = [f'{style}合计'] + [None] * 6
        for i in range(gn):
            cl = get_column_letter(8 + i)
            sub.append(f'=SUM({cl}{data_start}:{cl}{data_end})')
        for col in (total_col, loss_col, incl_col):
            cl = get_column_letter(col)
            sub.append(f'=SUM({cl}{data_start}:{cl}{data_end})')
        ws.append(sub)

        sub_row = ws.max_row
        ws.row_dimensions[sub_row].height = 22
        _style_row(ws, sub_row, n_cols, font=_f(11), alignment=ALIGN_C, border=BORDER_THIN)
        ws.cell(row=sub_row, column=1).font = _f(11, bold=True)
        ws.merge_cells(start_row=sub_row, start_column=1,
                       end_row=sub_row, end_column=7)

        # ── Spacer ────────────────────────────────────────────
        ws.append([None] * n_cols)
        ws.row_dimensions[ws.max_row].height = 8


# ══════════════════════════════════════════════════════════════
# 价格牌汇总 & 其他包装辅料汇总
# ══════════════════════════════════════════════════════════════

# Fills for the two new sheets
FILL_SECTION_HDR = PatternFill('solid', fgColor='D9E1F2')   # light blue – section heading
FILL_COL_HDR     = PatternFill('solid', fgColor='BDD7EE')   # medium blue – column heading
FILL_COLOUR_HDR  = PatternFill('solid', fgColor='E2EFDA')   # light green – colour group heading
FILL_SUBTOTAL2   = PatternFill('solid', fgColor='FFF2CC')   # yellow – 合计 / 含损耗
FILL_TOTAL_ALL   = PatternFill('solid', fgColor='FCE4D6')   # peach – 汇总(全部港口)
FILL_ACCESSORY   = PatternFill('solid', fgColor='F2F2F2')   # grey – accessory rows
FILL_GRP_CLOTH   = PatternFill('solid', fgColor='DEEAF1')   # light blue – clothing group header
FILL_GRP_NUMER   = PatternFill('solid', fgColor='E2EFDA')   # light green – numeric/shoe group header
CLOTHING_GRP_HDR = '衣码 (2XS~2XL)'
NUMERIC_GRP_HDR  = '特殊码 (10A~14D)'


def _is_xinma_port(port: str) -> bool:
    """Return True if port is a Singapore / Malaysia destination."""
    pu = str(port or '').upper()
    return any(x in pu for x in ('SINGAPORE', 'MALAYSIA', '新加坡', '马来西亚', '新马'))


def _is_brazil_port_name(port: str) -> bool:
    """Return True if port is a Brazil destination."""
    pu = str(port or '').upper()
    return any(x in pu for x in ('BRAZIL', '巴西'))


def _style_colour_str(rec) -> str:
    """Return 'STYLE-COLOUR' label (e.g. '6335908-02') for a record."""
    cs = str(rec.get('客款号', '') or '').strip()
    cc = str(rec.get('客户色号', '') or '').strip()
    if cs and cc:
        return f'{cs}-{cc}'
    return cs or cc or ''


def _get_all_sizes(records) -> list:
    """Return globally-sorted list of all size names present in records."""
    seen: list = []
    for rec in records:
        for s in rec.get('sizes_by_name', {}):
            if s not in seen:
                seen.append(s)
    seen.sort(key=_size_sort_key)
    return seen


def _split_sizes(sizes):
    """Split a sorted size list into (clothing_sizes, numeric_sizes).
    clothing = standard letter sizes (XS/S/M/L/XL/XXL/2XL etc.) – sort-key group 0.
    numeric  = sizes with digits (10A, 12B, 14D etc.) – sort-key group 1+.
    """
    clothing = [s for s in sizes if _size_sort_key(s)[0] == 0]
    numeric  = [s for s in sizes if _size_sort_key(s)[0] != 0]
    return clothing, numeric


# Full standard size sets — always shown in full when the group is non-empty.
# Any extra sizes in the data that fall outside these lists are appended after sorting.
_CLOTHING_STANDARD: list = ['2XS', 'XS', 'S', 'M', 'L', 'XL', '2XL']
_NUMERIC_STANDARD:  list = [
    '10A', '10B', '10C', '10D',
    '12A', '12B', '12C', '12D',
    '14A', '14B', '14C', '14D',
]


def _add_fixed_endpoints(sizes):
    """Expand each size group to its full standard set when that group is non-empty.

    If the data contains any clothing size → show all of 2XS/XS/S/M/L/XL/2XL
    (plus any data sizes outside that standard list, e.g. 3XL).
    If the data contains any numeric size  → show all of 10A-14D
    (plus any data sizes outside that range).
    Sizes with 0 quantity appear as empty cells.
    """
    clothing, numeric = _split_sizes(sizes)
    if clothing:
        merged = list(set(clothing) | set(_CLOTHING_STANDARD))
        merged.sort(key=_size_sort_key)
        clothing = merged
    if numeric:
        merged = list(set(numeric) | set(_NUMERIC_STANDARD))
        merged.sort(key=_size_sort_key)
        numeric = merged
    return clothing + numeric


def _write_grp_header_row(ws, fixed_n, clothing_sizes, numeric_sizes, total_n):
    """Append one group-label row spanning the size columns.

    fixed_n        – number of fixed (non-size) columns at the left
    clothing_sizes – list of clothing size names (may be empty)
    numeric_sizes  – list of numeric size names (may be empty)
    total_n        – total column count for this block (used to pad the row)

    Both group labels are always written (even if a group has 0 sizes).
    An empty group gets a 1-column slot for its label.
    Returns the row number just written.
    """
    c_n = len(clothing_sizes)
    n_n = len(numeric_sizes)
    # Build row: blanks over fixed cols + clothing label (1+ cols) + numeric label (1+ cols)
    row = [None] * fixed_n
    row.append(CLOTHING_GRP_HDR)
    row.extend([None] * max(c_n - 1, 0))
    row.append(NUMERIC_GRP_HDR)
    row.extend([None] * max(n_n - 1, 0))
    # Pad out to total_n with None
    row.extend([None] * max(total_n - len(row), 0))
    ws.append(row)
    grp_r = ws.max_row
    ws.row_dimensions[grp_r].height = 14
    # Merge clothing group
    c_start = fixed_n + 1
    c_end   = fixed_n + max(c_n, 1)
    if c_end > c_start:
        ws.merge_cells(start_row=grp_r, start_column=c_start, end_row=grp_r, end_column=c_end)
    # Merge numeric group
    n_start = c_end + 1
    n_end   = n_start + max(n_n, 1) - 1
    if n_end > n_start:
        ws.merge_cells(start_row=grp_r, start_column=n_start, end_row=grp_r, end_column=n_end)
    # Style clothing cells
    for c in range(c_start, c_end + 1):
        _apply(ws.cell(grp_r, c), font=_f(9, bold=True), alignment=ALIGN_C, fill=FILL_GRP_CLOTH)
    # Style numeric cells
    for c in range(n_start, n_end + 1):
        _apply(ws.cell(grp_r, c), font=_f(9, bold=True), alignment=ALIGN_C, fill=FILL_GRP_NUMER)
    return grp_r


def _sum_sizes_for(records, sizes) -> dict:
    """Sum each size column across a list of records. Returns {size: total}."""
    totals = {s: 0 for s in sizes}
    for rec in records:
        szm = rec.get('sizes_by_name', {})
        for s in sizes:
            v = szm.get(s)
            if isinstance(v, (int, float)):
                totals[s] += v
    return totals


def _with_loss_per_size(totals: dict, sizes: list) -> dict:
    """Apply _calc_loss per-size (each size quantity gets its own loss tier).
    Returns {size: qty_with_loss}.
    """
    result = {}
    for s in sizes:
        qty = totals.get(s, 0) or 0
        result[s] = (qty + _calc_loss(qty)) if qty > 0 else 0
    return result


def _jiage_pai_section(port: str) -> str:
    """Classify port into section heading for 价格牌汇总 (legacy, used when no config file)."""
    port_s = str(port or '').strip()
    if not port_s:
        return '未知港口'
    if _is_xinma_port(port_s):
        return '新马港口'
    for tab_name, cfg in PORT_TAB_CONFIG.items():
        if tab_name != '巴西' and port_s in cfg.get('ports', []):
            return '特殊港口'
    return port_s


# ── Config-file loaders ───────────────────────────────────────────────────────

def load_port_config_jiage(path):
    """Parse the 价格牌 sheet of CO港口分类.xlsx.

    Returns dict:
        'groups'     → OrderedDict {group_name: [dest_upper, ...]}  (column order)
        'thai_ports' → set of uppercased destination names (泰文 column)
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    if '价格牌' not in wb.sheetnames:
        return {'groups': OrderedDict(), 'thai_ports': set()}
    ws = wb['价格牌']
    n_cols = ws.max_column or 1
    headers = [ws.cell(1, c).value for c in range(1, n_cols + 1)]
    groups: OrderedDict = OrderedDict()
    thai_ports: set = set()
    for col_idx, header in enumerate(headers, start=1):
        if not header:
            continue
        h = str(header).strip()
        is_thai = '泰文' in h
        ports = []
        for row in range(2, (ws.max_row or 1) + 1):
            v = ws.cell(row, col_idx).value
            if v and str(v).strip():
                ports.append(str(v).strip().upper())
        if is_thai:
            thai_ports.update(ports)
        elif h:
            groups[h] = ports
    return {'groups': groups, 'thai_ports': thai_ports}


def load_port_config_qita(path):
    """Parse the 其他 sheet of CO港口分类.xlsx.

    Returns dict:
        'zuk_ports'    → set of uppercased destination names (特殊 column)
        'brazil_ports' → set of uppercased destination names (巴西 / BRAZIL column)
    """
    wb = openpyxl.load_workbook(path, data_only=True)
    if '其他' not in wb.sheetnames:
        return {'zuk_ports': set(), 'brazil_ports': set()}
    ws = wb['其他']
    n_cols = ws.max_column or 1
    headers = [ws.cell(1, c).value for c in range(1, n_cols + 1)]
    zuk_ports: set = set()
    brazil_ports: set = set()
    for col_idx, header in enumerate(headers, start=1):
        if not header:
            continue
        h = str(header).strip().upper()
        ports = []
        for row in range(2, (ws.max_row or 1) + 1):
            v = ws.cell(row, col_idx).value
            if v and str(v).strip():
                ports.append(str(v).strip().upper())
        if '特殊' in h:
            zuk_ports.update(ports)
        elif '巴西' in h or 'BRAZIL' in h:
            brazil_ports.update(ports)
    return {'zuk_ports': zuk_ports, 'brazil_ports': brazil_ports}


# ── Excel formula helpers ─────────────────────────────────────────────────────

def _loss_formula(cell_ref: str) -> str:
    """Return an Excel formula that applies the wastage tier to cell_ref.

    Tiers (matching _calc_loss):
        ≤0     → ""  (blank)
        1–10   → n+3
        11–30  → n+6
        31–100 → n+8
        101–300→ n+10
        301–500→ n+15
        501–800→ n+20
        801–1000→ n+25
        >1000  → n+CEILING(n*0.03, 1)
    """
    n = cell_ref
    return (
        f'=IF({n}<=0,"",IF({n}<=10,{n}+3,'
        f'IF({n}<=30,{n}+6,'
        f'IF({n}<=100,{n}+8,'
        f'IF({n}<=300,{n}+10,'
        f'IF({n}<=500,{n}+15,'
        f'IF({n}<=800,{n}+20,'
        f'IF({n}<=1000,{n}+25,'
        f'{n}+CEILING({n}*0.03,1)))))))))'
    )


def _pct_formula(total_ref: str, loss_ref: str) -> str:
    """损耗% formula: ROUND((loss-total)/total*100, 0); blank when total=0."""
    return (
        f'=IF({total_ref}>0,'
        f'ROUND(({loss_ref}-{total_ref})/{total_ref}*100,0),"")'
    )


def gen_jiage_pai(dest_wb, port_records_a, port_cfg_jiage=None):
    """Add 价格牌汇总 sheet to dest_wb.

    If port_cfg_jiage is provided (from load_port_config_jiage), groups are driven
    by the 价格牌 sheet of CO港口分类.xlsx:
      - Normal columns → named groups (e.g. '特殊港口', '新马港口')
      - 泰文 column     → Thai-tag overlay section appended at the end
    Unrecognised destinations become standalone groups labelled by their own name.

    Without a config, falls back to the legacy _jiage_pai_section() logic.

    Per colour block: column header | data rows | 合计 (SUM) | 含损耗 (formula) |
                      损耗% (formula) | 4 blank spacer rows.
    """
    ws = dest_wb.create_sheet(title='价格牌汇总')

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 14   # 厂款号
    ws.column_dimensions['D'].width = 20   # Destination
    for i in range(30):
        ws.column_dimensions[get_column_letter(5 + i)].width = 7

    # ── Build section → records map ────────────────────────────────────────
    if port_cfg_jiage:
        groups     = port_cfg_jiage['groups']      # OrderedDict {name: [dest_up]}
        thai_ports = port_cfg_jiage['thai_ports']  # set of uppercase dests
        dest_to_group: dict = {}
        for gn, ports in groups.items():
            for p in ports:
                dest_to_group[p] = gn            # keys already uppercase
        group_order = list(groups.keys())
    else:
        # legacy path: use _jiage_pai_section
        groups        = OrderedDict()
        thai_ports    = set()
        dest_to_group = {}
        group_order   = []

    grouped_recs: dict = {gn: [] for gn in group_order}
    standalone_order: list = []          # first-occurrence dest (raw string)
    standalone_recs:  dict = {}          # dest_upper → [records]

    for rec in port_records_a:
        # Use raw English destination name for config matching (substring, case-insensitive)
        dest_raw = str(rec.get('destination_raw', '') or '').strip().upper()
        dest_cn  = str(rec.get('港口', '') or '').strip()
        dest_for_match = dest_raw or dest_cn.upper()   # fall back to Chinese name if no raw

        if port_cfg_jiage:
            # Substring match: config port string somewhere in the raw destination name
            matched_group = None
            for cfg_up, gn in dest_to_group.items():
                if cfg_up in dest_for_match:
                    matched_group = gn
                    break
            if matched_group:
                grouped_recs[matched_group].append(rec)
            else:
                # Standalone (unrecognised by config)
                dest_disp = dest_cn or dest_raw  # prefer Chinese display name
                key = dest_disp.upper()
                if key not in standalone_recs:
                    standalone_order.append(dest_disp)
                    standalone_recs[key] = []
                standalone_recs[key].append(rec)
        else:
            # legacy grouping
            sec = _jiage_pai_section(dest_cn)
            if sec not in grouped_recs:
                grouped_recs[sec] = []
                group_order.append(sec)
            grouped_recs[sec].append(rec)

    # ── Shared inner writer ─────────────────────────────────────────────────
    def _write_colour_block(sc, grp):
        """Write col-header + data + 合计 + 含损耗 + 损耗% + 4 blanks."""
        colour_sizes = _add_fixed_endpoints(_get_all_sizes(grp))
        fixed_n = 4   # PO_NUMBER | STYLE_COLOUR | 厂款号 | Destination
        n = fixed_n + len(colour_sizes)

        # ── Column header ────────────────────────────────────────────────────
        ws.append(['PO_NUMBER', 'STYLE_COLOUR', '厂款号', 'Destination'] + colour_sizes)
        hdr_r = ws.max_row
        ws.row_dimensions[hdr_r].height = 18
        for c in range(1, n + 1):
            _apply(ws.cell(hdr_r, c), font=_f(10, bold=True),
                   alignment=ALIGN_C, fill=FILL_COL_HDR, border=BORDER_THIN)

        # ── Data rows ────────────────────────────────────────────────────────
        first_data_r = ws.max_row + 1
        for rec in grp:
            po         = str(rec.get('合同号：', '') or rec.get('合同号', '') or '').strip()
            factory_no = str(rec.get('厂款号：', '') or rec.get('厂款号', '') or '').strip()
            dest       = str(rec.get('港口', '') or '').strip()
            szm        = rec.get('sizes_by_name', {})
            ws.append([po, sc, factory_no, dest] + [
                szm.get(s) if isinstance(szm.get(s), (int, float)) and szm.get(s)
                else None
                for s in colour_sizes
            ])
            r = ws.max_row
            ws.row_dimensions[r].height = 18
            for c in range(1, n + 1):
                _apply(ws.cell(r, c), font=_f(10), alignment=ALIGN_C, border=BORDER_THIN)
        last_data_r = ws.max_row

        # ── 合计 – SUM formula per size column ──────────────────────────────
        total_r   = ws.max_row + 1
        total_row = [None, None, None, '合计']
        for ci in range(5, n + 1):
            cl = get_column_letter(ci)
            total_row.append(f'=SUM({cl}{first_data_r}:{cl}{last_data_r})')
        ws.append(total_row)
        ws.row_dimensions[total_r].height = 18
        for c in range(1, n + 1):
            _apply(ws.cell(total_r, c), font=_f(10, bold=(c == 4)),
                   alignment=ALIGN_C, fill=FILL_SUBTOTAL2, border=BORDER_THIN)

        # ── 含损耗 ────────────────────────────────────────────────────────────
        loss_r   = ws.max_row + 1
        loss_row = [None, None, None, '含损耗']
        for ci in range(5, n + 1):
            cl = get_column_letter(ci)
            loss_row.append(_loss_formula(f'{cl}{total_r}'))
        ws.append(loss_row)
        ws.row_dimensions[loss_r].height = 18
        for c in range(1, n + 1):
            _apply(ws.cell(loss_r, c), font=_f(10, bold=(c == 4)),
                   alignment=ALIGN_C, fill=FILL_SUBTOTAL2, border=BORDER_THIN)

        # ── 损耗% ────────────────────────────────────────────────────────────
        pct_r   = ws.max_row + 1
        pct_row = [None, None, None, '损耗%']
        for ci in range(5, n + 1):
            cl = get_column_letter(ci)
            pct_row.append(_pct_formula(f'{cl}{total_r}', f'{cl}{loss_r}'))
        ws.append(pct_row)
        ws.row_dimensions[pct_r].height = 18
        for c in range(1, n + 1):
            _apply(ws.cell(pct_r, c), font=_f(10, bold=(c == 4)),
                   alignment=ALIGN_C, fill=FILL_SUBTOTAL2, border=BORDER_THIN)

        # ── 4 blank spacer rows ───────────────────────────────────────────────
        for _ in range(4):
            ws.append([None] * n)
            ws.row_dimensions[ws.max_row].height = 8

    def _write_section(section_name, recs):
        """Write a labelled section containing per-colour colour blocks."""
        # Build colour groups (data order)
        colour_order_l: list = []
        colour_grps_l:  dict = {}
        for rec in recs:
            sc = _style_colour_str(rec)
            if sc not in colour_grps_l:
                colour_grps_l[sc] = []
                colour_order_l.append(sc)
            colour_grps_l[sc].append(rec)
        # Keep same 款号 (style before '-') adjacent: styles in first-seen order, colours sorted
        style_first: dict = {}
        for sc in colour_order_l:
            style_first.setdefault(sc.split('-')[0], len(style_first))
        colour_order_l.sort(key=lambda sc: (style_first[sc.split('-')[0]], sc))

        max_cols = max(
            (4 + len(_add_fixed_endpoints(_get_all_sizes(colour_grps_l[sc]))) for sc in colour_order_l),
            default=5,
        )
        # Section header
        ws.append([section_name] + [None] * (max_cols - 1))
        sec_row = ws.max_row
        ws.row_dimensions[sec_row].height = 22
        ws.merge_cells(start_row=sec_row, start_column=1,
                       end_row=sec_row,   end_column=max_cols)
        cell = ws.cell(sec_row, 1)
        cell.value = section_name; cell.font = _f(12, bold=True)
        cell.alignment = ALIGN_C; cell.fill = FILL_SECTION_HDR

        for sc in colour_order_l:
            _write_colour_block(sc, colour_grps_l[sc])

    # ── Write configured groups ────────────────────────────────────────────
    first_sec = True
    for gn in group_order:
        recs = grouped_recs.get(gn, [])
        if not recs:
            continue
        if not first_sec:
            ws.append([None] * 10)
            ws.row_dimensions[ws.max_row].height = 8
        _write_section(gn, recs)
        first_sec = False

    # ── Write standalone groups (unrecognised destinations) ───────────────
    for dest_orig in standalone_order:
        dest_up = dest_orig.upper()
        recs = standalone_recs.get(dest_up, [])
        if not recs:
            continue
        if not first_sec:
            ws.append([None] * 10)
            ws.row_dimensions[ws.max_row].height = 8
        _write_section(dest_orig, recs)
        first_sec = False

    # ── Thai-tag overlay section ──────────────────────────────────────────
    if thai_ports:
        thai_recs = [
            r for r in port_records_a
            if any(
                tp in (str(r.get('destination_raw','') or '').strip().upper()
                       or str(r.get('港口','') or '').strip().upper())
                for tp in thai_ports
            )
        ]
        if thai_recs:
            if not first_sec:
                ws.append([None] * 10)
                ws.row_dimensions[ws.max_row].height = 8
            # Thai section: all colours combined in one flat block
            all_thai_sizes = _add_fixed_endpoints(_get_all_sizes(thai_recs))
            nt = 4 + len(all_thai_sizes)   # 4 fixed cols
            # Section header
            ws.append(['泰文吊牌汇总'] + [None] * (nt - 1))
            sec_row = ws.max_row
            ws.row_dimensions[sec_row].height = 22
            ws.merge_cells(start_row=sec_row, start_column=1,
                           end_row=sec_row,   end_column=nt)
            cell = ws.cell(sec_row, 1)
            cell.value = '泰文吊牌汇总'; cell.font = _f(12, bold=True)
            cell.alignment = ALIGN_C; cell.fill = FILL_SECTION_HDR
            # Column header
            ws.append(['PO_NUMBER', 'STYLE_COLOUR', '厂款号', 'Destination'] + all_thai_sizes)
            hdr_r = ws.max_row
            ws.row_dimensions[hdr_r].height = 18
            for c in range(1, nt + 1):
                _apply(ws.cell(hdr_r, c), font=_f(10, bold=True),
                       alignment=ALIGN_C, fill=FILL_COL_HDR, border=BORDER_THIN)
            # Detail rows (all colours together)
            first_thai_r = ws.max_row + 1
            for rec in thai_recs:
                po         = str(rec.get('合同号：', '') or rec.get('合同号', '') or '').strip()
                sc         = _style_colour_str(rec)
                factory_no = str(rec.get('厂款号：', '') or rec.get('厂款号', '') or '').strip()
                dest       = str(rec.get('港口', '') or '').strip()
                szm        = rec.get('sizes_by_name', {})
                ws.append([po, sc, factory_no, dest] + [
                    szm.get(s) if isinstance(szm.get(s), (int, float)) and szm.get(s)
                    else None
                    for s in all_thai_sizes
                ])
                r = ws.max_row
                ws.row_dimensions[r].height = 18
                for c in range(1, nt + 1):
                    _apply(ws.cell(r, c), font=_f(10),
                           alignment=ALIGN_C, border=BORDER_THIN)
            last_thai_r = ws.max_row
            # 总数 (合并颜色) row
            total_r   = ws.max_row + 1
            total_row = [None, None, None, '总数 (合并颜色)']
            for ci in range(5, nt + 1):
                cl = get_column_letter(ci)
                total_row.append(f'=SUM({cl}{first_thai_r}:{cl}{last_thai_r})')
            ws.append(total_row)
            ws.row_dimensions[total_r].height = 18
            for c in range(1, nt + 1):
                _apply(ws.cell(total_r, c), font=_f(10, bold=(c == 4)),
                       alignment=ALIGN_C, fill=FILL_TOTAL_ALL, border=BORDER_THIN)
            # 含损耗 row
            loss_r   = ws.max_row + 1
            loss_row = [None, None, None, '含损耗']
            for ci in range(5, nt + 1):
                cl = get_column_letter(ci)
                loss_row.append(_loss_formula(f'{cl}{total_r}'))
            ws.append(loss_row)
            ws.row_dimensions[loss_r].height = 18
            for c in range(1, nt + 1):
                _apply(ws.cell(loss_r, c), font=_f(10, bold=(c == 4)),
                       alignment=ALIGN_C, fill=FILL_SUBTOTAL2, border=BORDER_THIN)


def gen_qita_fuliao(dest_wb, port_records_a, port_cfg_qita=None):
    """Add 其他包装辅料汇总 sheet to dest_wb.

    If port_cfg_qita is provided (from load_port_config_qita), ports are classified
    using the 其他 sheet of CO港口分类.xlsx:
        zuk    = destinations in the 特殊 column
        brazil = destinations in the 巴西/BRAZIL column
        normal = everything else

    Without a config, falls back to the legacy PORT_TAB_CONFIG classification.

    Section A – ZUK 洗水唛
        Per colour: colour header | col header | zuk data rows | 合计 | 含损耗 |
                    损耗% | blank row.

    Section B – 洗水唛 (其他港口 + 巴西)
        Per colour:
          colour header | col header |
          normal rows → 合计(其他港口) → 含损耗(其他港口) → 损耗%(其他港口) |
          [brazil rows → 合计(巴西) → 含损耗(巴西) → 损耗%(巴西)]  (if any) |
          5 cross-section summary lines:
              主唛 (全部港口)   = 含损耗_normal + 含损耗_brazil + 含损耗_zuk
              ZUK 洗水唛        = 含损耗_zuk
              洗水唛 (含巴西)   = 含损耗_normal + 含损耗_brazil
              尺码牌 (不含巴西) = 含损耗_normal + 含损耗_zuk
              巴西尺码牌        = 含损耗_brazil
          2 blank rows.
    """
    ws = dest_wb.create_sheet(title='其他包装辅料汇总')

    # Global sizes → used only for section-header merge widths
    _global_sizes = _get_all_sizes(port_records_a)
    global_n = 4 + max(len(_global_sizes), 1)

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 14
    ws.column_dimensions['C'].width = 14   # 厂款号
    ws.column_dimensions['D'].width = 20   # Destination
    for i in range(max(len(_global_sizes), 1)):
        ws.column_dimensions[get_column_letter(5 + i)].width = 7

    # ── Port classifier ────────────────────────────────────────────────────
    if port_cfg_qita:
        _zuk_ports    = port_cfg_qita['zuk_ports']    # set of uppercase strings
        _brazil_ports = port_cfg_qita['brazil_ports']
    else:
        # Legacy fallback: use PORT_TAB_CONFIG
        _zuk_ports:    set = set()
        _brazil_ports: set = set()
        for tab_name, cfg in PORT_TAB_CONFIG.items():
            if tab_name == '巴西':
                _brazil_ports.update(p.upper() for p in cfg.get('ports', []))
            else:
                _zuk_ports.update(p.upper() for p in cfg.get('ports', []))

    def _classify(dest_str: str) -> str:
        """Return 'zuk' | 'brazil' | 'normal' for a destination string."""
        pu = dest_str.strip().upper()
        if port_cfg_qita:
            if any(p in pu for p in _zuk_ports):
                return 'zuk'
            if any(p in pu for p in _brazil_ports):
                return 'brazil'
            return 'normal'
        else:
            # Legacy: use Chinese port name matching
            if any(p in pu for p in _zuk_ports):
                return 'zuk'
            if any(p in pu for p in _brazil_ports):
                return 'brazil'
            return 'normal'

    # ── Group records by STYLE_COLOUR ─────────────────────────────────────
    colour_order: list = []
    colour_grps:  dict = {}
    for rec in port_records_a:
        sc = _style_colour_str(rec)
        if sc not in colour_grps:
            colour_grps[sc] = []
            colour_order.append(sc)
        colour_grps[sc].append(rec)

    # ── Shared helpers ─────────────────────────────────────────────────────
    def _data_rows(recs, colour_sizes, n, sc_label):
        """Append data rows and return (first_row, last_row)."""
        first_r = ws.max_row + 1
        for rec in recs:
            po         = str(rec.get('合同号：', '') or rec.get('合同号', '') or '').strip()
            factory_no = str(rec.get('厂款号：', '') or rec.get('厂款号', '') or '').strip()
            dest       = str(rec.get('港口', '') or '').strip()
            szm        = rec.get('sizes_by_name', {})
            ws.append([po, sc_label, factory_no, dest] + [
                szm.get(s) if isinstance(szm.get(s), (int, float)) and szm.get(s)
                else None
                for s in colour_sizes
            ])
            r = ws.max_row
            ws.row_dimensions[r].height = 18
            for c in range(1, n + 1):
                _apply(ws.cell(r, c), font=_f(10), alignment=ALIGN_C, border=BORDER_THIN)
        return first_r, ws.max_row

    def _formula_row(label, formula_list, n, fill=FILL_SUBTOTAL2):
        """Append a summary row; formula_list has one entry per size column."""
        ws.append([None, None, None, label] + formula_list)
        r = ws.max_row
        ws.row_dimensions[r].height = 18
        for c in range(1, n + 1):
            _apply(ws.cell(r, c), font=_f(10, bold=(c == 4)),
                   alignment=ALIGN_C, fill=fill, border=BORDER_THIN)
        return r

    def _col_header(colour_sizes, n):
        ws.append(['PO_NUMBER', 'STYLE_COLOUR', '厂款号', 'Destination'] + colour_sizes)
        ch = ws.max_row
        ws.row_dimensions[ch].height = 18
        for c in range(1, n + 1):
            _apply(ws.cell(ch, c), font=_f(10, bold=True),
                   alignment=ALIGN_C, fill=FILL_COL_HDR, border=BORDER_THIN)

    def _colour_header(sc, n):
        ws.append([f'颜色 {sc}'] + [None] * (n - 1))
        cr = ws.max_row
        ws.row_dimensions[cr].height = 20
        ws.merge_cells(start_row=cr, start_column=1, end_row=cr, end_column=n)
        cell = ws.cell(cr, 1)
        cell.value = f'颜色 {sc}'; cell.font = _f(11, bold=True)
        cell.alignment = ALIGN_C;  cell.fill = FILL_COLOUR_HDR

    # ── SECTION A: ZUK 洗水唛 ─────────────────────────────────────────────
    zuk_port_names = sorted(
        str(rec.get('港口', '') or '').strip()
        for rec in port_records_a
        if _classify(str(rec.get('destination_raw', '') or '') or str(rec.get('港口', '') or '')) == 'zuk'
    )
    _uniq_zuk = list(dict.fromkeys(zuk_port_names))
    sa_title = 'ZUK 洗水唛 (适用港口: ' + ' + '.join(_uniq_zuk[:4]) + ')' if _uniq_zuk \
               else 'ZUK 洗水唛'

    ws.append([sa_title] + [None] * (global_n - 1))
    sa_hdr = ws.max_row
    ws.row_dimensions[sa_hdr].height = 24
    ws.merge_cells(start_row=sa_hdr, start_column=1, end_row=sa_hdr, end_column=global_n)
    cell = ws.cell(sa_hdr, 1)
    cell.font = _f(12, bold=True); cell.alignment = ALIGN_C; cell.fill = FILL_SECTION_HDR

    # Track per-colour ZUK 含损耗 row numbers + col maps for cross-references
    zuk_loss_info: dict = {}  # sc → {'loss_row': int, 'col_map': {size: col_letter}}

    for sc in colour_order:
        zuk_recs = [r for r in colour_grps[sc]
                    if _classify(str(r.get('destination_raw', '') or '') or str(r.get('港口', '') or '')) == 'zuk']
        if not zuk_recs:
            continue
        colour_sizes = _add_fixed_endpoints(_get_all_sizes(zuk_recs))
        n = 4 + len(colour_sizes)
        col_map = {s: get_column_letter(5 + i) for i, s in enumerate(colour_sizes)}

        _colour_header(sc, n)
        _col_header(colour_sizes, n)

        first_r, last_r = _data_rows(zuk_recs, colour_sizes, n, sc)

        # 合计
        total_r = ws.max_row + 1
        _formula_row('合计',
                     [f'=SUM({get_column_letter(ci)}{first_r}:{get_column_letter(ci)}{last_r})'
                      for ci in range(5, n + 1)], n)

        # 含损耗
        loss_r = ws.max_row + 1
        _formula_row('含损耗',
                     [_loss_formula(f'{get_column_letter(ci)}{total_r}')
                      for ci in range(5, n + 1)], n)

        # 损耗%
        _formula_row('损耗%',
                     [_pct_formula(f'{get_column_letter(ci)}{total_r}',
                                   f'{get_column_letter(ci)}{loss_r}')
                      for ci in range(5, n + 1)], n)

        zuk_loss_info[sc] = {'loss_row': loss_r, 'col_map': col_map}

        ws.append([None] * n)
        ws.row_dimensions[ws.max_row].height = 8

    # ── 5 blank rows between sections ─────────────────────────────────────
    for _ in range(5):
        ws.append([None] * global_n)
        ws.row_dimensions[ws.max_row].height = 8

    # ── SECTION B: 洗水唛 (其他港口 + 巴西) ──────────────────────────────
    ws.append(['洗水唛 (其他港口 + 巴西)'] + [None] * (global_n - 1))
    sb_hdr = ws.max_row
    ws.row_dimensions[sb_hdr].height = 24
    ws.merge_cells(start_row=sb_hdr, start_column=1, end_row=sb_hdr, end_column=global_n)
    cell = ws.cell(sb_hdr, 1)
    cell.font = _f(12, bold=True); cell.alignment = ALIGN_C; cell.fill = FILL_SECTION_HDR

    for sc in colour_order:
        normal_recs = [r for r in colour_grps[sc]
                       if _classify(str(r.get('destination_raw', '') or '') or str(r.get('港口', '') or '')) == 'normal']
        brazil_recs = [r for r in colour_grps[sc]
                       if _classify(str(r.get('destination_raw', '') or '') or str(r.get('港口', '') or '')) == 'brazil']
        if not normal_recs and not brazil_recs:
            continue

        # Column layout covers ALL non-ZUK records for this colour
        b_recs       = normal_recs + brazil_recs
        colour_sizes = _add_fixed_endpoints(_get_all_sizes(b_recs))
        n            = 4 + len(colour_sizes)
        b_col_map    = {s: get_column_letter(5 + i) for i, s in enumerate(colour_sizes)}

        _colour_header(sc, n)
        _col_header(colour_sizes, n)

        # ── Normal rows + 合计/含损耗/损耗% ────────────────────────────
        if normal_recs:
            fn, ln = _data_rows(normal_recs, colour_sizes, n, sc)
        else:
            fn = ln = ws.max_row  # no rows; formulas will sum empty range

        normal_total_r = ws.max_row + 1
        _formula_row('合计 (其他港口)',
                     [f'=SUM({get_column_letter(ci)}{fn}:{get_column_letter(ci)}{ln})'
                      if normal_recs else '=0'
                      for ci in range(5, n + 1)], n)

        normal_loss_r = ws.max_row + 1
        _formula_row('含损耗 (其他港口)',
                     [_loss_formula(f'{get_column_letter(ci)}{normal_total_r}')
                      for ci in range(5, n + 1)], n)

        _formula_row('损耗% (其他港口)',
                     [_pct_formula(f'{get_column_letter(ci)}{normal_total_r}',
                                   f'{get_column_letter(ci)}{normal_loss_r}')
                      for ci in range(5, n + 1)], n)

        # ── Brazil rows + 合计/含损耗/损耗% (if any) ────────────────────
        brazil_loss_r = None
        if brazil_recs:
            fb, lb = _data_rows(brazil_recs, colour_sizes, n, sc)

            brazil_total_r = ws.max_row + 1
            _formula_row('合计 (巴西)',
                         [f'=SUM({get_column_letter(ci)}{fb}:{get_column_letter(ci)}{lb})'
                          for ci in range(5, n + 1)], n)

            brazil_loss_r = ws.max_row + 1
            _formula_row('含损耗 (巴西)',
                         [_loss_formula(f'{get_column_letter(ci)}{brazil_total_r}')
                          for ci in range(5, n + 1)], n)

            _formula_row('损耗% (巴西)',
                         [_pct_formula(f'{get_column_letter(ci)}{brazil_total_r}',
                                       f'{get_column_letter(ci)}{brazil_loss_r}')
                          for ci in range(5, n + 1)], n)

        # ── 5 cross-section summary lines ────────────────────────────────
        zuk_info = zuk_loss_info.get(sc)

        SUMMARY_DEFS = [
            ('主唛 (全部港口)',    ['normal', 'brazil', 'zuk']),
            ('ZUK 洗水唛',        ['zuk']),
            ('洗水唛 (含巴西)',    ['normal', 'brazil']),
            ('尺码牌 (不含巴西)', ['normal', 'zuk']),
            ('巴西尺码牌',        ['brazil']),
        ]
        for label, components in SUMMARY_DEFS:
            formulas = []
            for s in colour_sizes:
                b_cl  = b_col_map[s]
                terms = []
                if 'zuk' in components and zuk_info and s in zuk_info['col_map']:
                    zuk_cl = zuk_info['col_map'][s]
                    terms.append(f'{zuk_cl}{zuk_info["loss_row"]}')
                if 'normal' in components:
                    terms.append(f'{b_cl}{normal_loss_r}')
                if 'brazil' in components and brazil_loss_r:
                    terms.append(f'{b_cl}{brazil_loss_r}')
                formulas.append(('=' + '+'.join(f'N({t})' for t in terms)) if terms else '')
            _formula_row(label, formulas, n, fill=FILL_ACCESSORY)

        # 2 blank rows
        for _ in range(2):
            ws.append([None] * n)
            ws.row_dimensions[ws.max_row].height = 8


# 特殊港口打印单的字号与行高（如需调整打印大小，改这里即可）
SLIP_TITLE_SIZE,  SLIP_TITLE_H  = 22, 34
SLIP_HEADER_SIZE, SLIP_HEADER_H = 16, 30
SLIP_DATA_SIZE,   SLIP_DATA_H   = 16, 46
SLIP_NOTES_SIZE,  SLIP_NOTES_H  = 13, 110
SLIP_FOOTER_SIZE, SLIP_FOOTER_H = 16, 30

# 「单件装」打印单备注（{port}=该片去重港口，{style}=该片客款号）
SINGLE_SLIP_NOTE = (
    '注:\n'
    '1：{port}港口,洗水唛不一样 客款号是{style} ,注意与大货区分出来,'
    '回货单上面注明特殊港口单独回货，并且一色齐码齐数后查好再送回\n'
    '2：此单一式两份，放一份在货里随货送回公司\n'
    '3：特殊港口的货，请工厂物必重视，按要求回货，否则一次警告，二次罚款200元/次'
)


def gen_special_port(dest_wb, port_records, tab_name, year_quarter, cfg):
    """Generate 巴西 / 英国诺丁汉 / 特殊港口 (print-ready, one slip per page).
    Block layout per record (sizes/heights from SLIP_* constants above):
      title  (merged)
      header (borders)
      data   (borders)
      notes  (merged, left-aligned)
      footer (merged)
      2 blank rows
    A manual page break after each block makes each slip print on its own
    A4 landscape page (fit-to-width).
    """
    ws = dest_wb.create_sheet(title=tab_name)
    port_set  = set(cfg['ports'])
    filtered  = [r for r in port_records if r['港口'] in port_set]

    if not filtered:
        return

    # ── Group lines sharing (客款号, 厂款号, 颜色), preserving encounter order ──
    groups: "OrderedDict[tuple, list]" = OrderedDict()
    for rec in filtered:
        key = (rec['客款号'], rec['厂款号'], rec['颜色'])
        groups.setdefault(key, []).append(rec)

    # Per-group merged size header (union of the group's size_headers)
    def _group_sh(grp):
        raw: list = []
        for rec in grp:
            for s in rec['size_headers']:
                if s not in raw:
                    raw.append(s)
        return _add_fixed_endpoints(raw)

    group_sh = {key: _group_sh(grp) for key, grp in groups.items()}
    max_n = max(len(sh) for sh in group_sh.values())

    # ── Column widths (use max_n so date/notes columns are always sized) ──
    fixed_widths = {1: 16, 2: 13, 3: 13, 4: 16, 5: 13, 6: 20, 7: 16}  # 合同号…港口
    for c, w in fixed_widths.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    for c in range(8, 8 + max_n):                        # 尺码带
        ws.column_dimensions[get_column_letter(c)].width = 8.5
    ws.column_dimensions[get_column_letter(8 + max_n)].width  = 11      # 总数
    ws.column_dimensions[get_column_letter(9 + max_n)].width  = 19.73   # 工厂出厂期
    ws.column_dimensions[get_column_letter(10 + max_n)].width = 30      # 订单附加备注

    block_end_rows = []   # last row of each block → page-break anchors

    for key, grp in groups.items():
        sh      = group_sh[key]
        n       = len(sh)
        n_cols  = 10 + n   # total data columns for this block

        # ── Title ────────────────────────────────────────────
        ws.append([f'{tab_name} {year_quarter}'])
        r_title = ws.max_row
        ws.row_dimensions[r_title].height = SLIP_TITLE_H
        ws.cell(row=r_title, column=1).value = f'{tab_name} {year_quarter}'
        _apply(ws.cell(row=r_title, column=1), font=_f(SLIP_TITLE_SIZE), alignment=ALIGN_C)
        ws.merge_cells(start_row=r_title, start_column=1,
                       end_row=r_title, end_column=n_cols)

        # ── Header ───────────────────────────────────────────
        ws.append(
            ['合同号：', '客款号', '厂款号：', '颜色', '客户色号', '英文颜色', '港口']
            + sh + ['总数', '工厂出厂期', '订单附加备注']
        )
        r_hdr = ws.max_row
        ws.row_dimensions[r_hdr].height = SLIP_HEADER_H
        _style_row(ws, r_hdr, n_cols, font=_f(SLIP_HEADER_SIZE), alignment=ALIGN_C, border=BORDER_THIN)

        # ── Data lines (one per record in the group) ─────────
        r_first = ws.max_row + 1
        for rec in grp:
            sz_map    = dict(zip(rec['size_headers'], rec['sizes']))
            size_vals = [sz_map.get(s, None) for s in sh]
            ws.append(
                [rec['合同号'], rec['客款号'], rec['厂款号'], rec['颜色'],
                 rec['客户色号'], rec['英文颜色'], rec['港口']]
                + size_vals
                + [rec['总数'], rec['工厂出厂期'], rec['订单附加备注']]
            )
            r_data = ws.max_row
            ws.row_dimensions[r_data].height = SLIP_DATA_H
            _style_row(ws, r_data, n_cols, font=_f(SLIP_DATA_SIZE),
                       alignment=ALIGN_C, border=BORDER_THIN)
        r_last = ws.max_row

        # ── Summary line (合计: per-size + 总数 sums over the group's rows) ──
        sum_row: list = ['合计'] + [None] * 6            # merged over cols 1-7
        for ci in range(8, 8 + n):                       # size columns
            cl = get_column_letter(ci)
            sum_row.append(f'=SUM({cl}{r_first}:{cl}{r_last})')
        cl_total = get_column_letter(8 + n)              # 总数 column
        sum_row.append(f'=SUM({cl_total}{r_first}:{cl_total}{r_last})')
        sum_row += [None, None]                          # 工厂出厂期 / 订单附加备注
        ws.append(sum_row)
        r_sum = ws.max_row
        ws.row_dimensions[r_sum].height = SLIP_DATA_H
        _style_row(ws, r_sum, n_cols, font=_f(SLIP_DATA_SIZE, bold=True),
                   alignment=ALIGN_C, fill=FILL_SUBTOTAL, border=BORDER_THIN)
        ws.merge_cells(start_row=r_sum, start_column=1, end_row=r_sum, end_column=7)

        # ── Notes（组内去重港口以 / 连接填入 {port}）────────────
        grp_ports = []
        for rec in grp:
            if rec['港口'] and rec['港口'] not in grp_ports:
                grp_ports.append(rec['港口'])
        _note_text = cfg['notes'].format(port='/'.join(grp_ports))
        ws.append([_note_text])
        r_notes = ws.max_row
        ws.row_dimensions[r_notes].height = SLIP_NOTES_H
        ws.cell(row=r_notes, column=1).value = _note_text
        _apply(ws.cell(row=r_notes, column=1), font=_f(SLIP_NOTES_SIZE), alignment=ALIGN_L)
        ws.merge_cells(start_row=r_notes, start_column=1,
                       end_row=r_notes, end_column=n_cols)

        # ── Footer ───────────────────────────────────────────
        ws.append(['此单随货送回公司'])
        r_foot = ws.max_row
        ws.row_dimensions[r_foot].height = SLIP_FOOTER_H
        ws.cell(row=r_foot, column=1).value = '此单随货送回公司'
        _apply(ws.cell(row=r_foot, column=1), font=_f(SLIP_FOOTER_SIZE), alignment=ALIGN_C)
        ws.merge_cells(start_row=r_foot, start_column=1,
                       end_row=r_foot, end_column=n_cols)

        # ── Blank rows ───────────────────────────────────────
        ws.append([None])
        ws.append([None])
        block_end_rows.append(ws.max_row)   # end of this block

    # ── Page breaks: one group-slip per page (skip after final block) ──
    for r in block_end_rows[:-1]:
        ws.row_breaks.append(Break(id=r))

    # ── Print setup: A4 landscape, fit to width, centered ────────
    last_col = get_column_letter(10 + max_n)
    ws.print_area = f'A1:{last_col}{ws.max_row}'
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize   = 9            # A4
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0            # manual breaks govern vertical pages
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top  = ws.page_margins.bottom = 0.5


def gen_single_piece_slips(dest_wb, port_records, year_quarter):
    """「单件装」打印页：仅单件装(非套装)记录，按 (客款号,厂款号,颜色,港口类别) 每片一页。

    版式同 gen_special_port（标题/表头/数据/合计/注:/页脚，A4 横向、每片分页），
    片内行像「颜色套装港口汇总」合并：同 (客户色号,内部色号,英文颜色,工厂出厂期) 的
    合同号/港口 各 / 连接、尺码求和。备注块用 SINGLE_SLIP_NOTE（{port}=去重港口、{style}=客款号）。
    """
    singles = [r for r in port_records
               if _piece_label(r.get('订单附加备注')) == '单件装']
    if not singles:
        return

    ws = dest_wb.create_sheet(title='单件装')

    # ── 分片：(客款号, 厂款号, 颜色, 港口类别)，首现序 ──
    slips: "OrderedDict[tuple, list]" = OrderedDict()
    for rec in singles:
        skey = (rec['客款号'], rec['厂款号'], rec['颜色'], _port_class(rec['港口']))
        slips.setdefault(skey, []).append(rec)

    # ── 片内合并：同 (客户色号,内部色号,英文颜色,工厂出厂期) → 合同号/港口 / 连接、尺码求和 ──
    def _merge_lines(grp):
        order, merged = [], {}
        for r in grp:
            lkey = (r['客户色号'], r['内部色号'], r['英文颜色'], str(r['工厂出厂期'] or ''))
            if lkey not in merged:
                merged[lkey] = []
                order.append(lkey)
            merged[lkey].append(r)
        lines = []
        for lkey in order:
            rows = merged[lkey]
            pos, ports, acc, total = [], [], {}, 0
            for r in rows:
                po = str(r['合同号'] or '').strip()
                if po and po not in pos:
                    pos.append(po)
                p = str(r['港口'] or '').strip()
                if p and p not in ports:
                    ports.append(p)
                for s, v in zip(r['size_headers'], r['sizes']):
                    if isinstance(v, (int, float)):
                        acc[s] = acc.get(s, 0) + v
                t = r['总数']
                if isinstance(t, (int, float)):
                    total += t
            g0 = rows[0]
            lines.append({
                '合同号': '/'.join(pos), '客款号': g0['客款号'], '厂款号': g0['厂款号'],
                '颜色': g0['颜色'], '客户色号': g0['客户色号'], '英文颜色': g0['英文颜色'],
                '港口': '/'.join(ports), 'sizes': acc, '总数': total or None,
                '工厂出厂期': g0['工厂出厂期'], 'ports': ports})
        return lines

    slip_lines = {k: _merge_lines(g) for k, g in slips.items()}

    def _slip_sh(lines):
        raw: list = []
        for ln in lines:
            for s in ln['sizes']:
                if s not in raw:
                    raw.append(s)
        raw.sort(key=_size_sort_key)
        return _add_fixed_endpoints(raw)

    slip_sh = {k: _slip_sh(lines) for k, lines in slip_lines.items()}
    max_n = max((len(sh) for sh in slip_sh.values()), default=7)

    # ── 列宽（同 gen_special_port）──
    for c, w in {1: 16, 2: 13, 3: 13, 4: 16, 5: 13, 6: 20, 7: 16}.items():
        ws.column_dimensions[get_column_letter(c)].width = w
    for c in range(8, 8 + max_n):
        ws.column_dimensions[get_column_letter(c)].width = 8.5
    ws.column_dimensions[get_column_letter(8 + max_n)].width  = 11
    ws.column_dimensions[get_column_letter(9 + max_n)].width  = 19.73
    ws.column_dimensions[get_column_letter(10 + max_n)].width = 30

    block_end_rows = []

    for skey, lines in slip_lines.items():
        client_style, factory, colour, _pcls = skey
        sh = slip_sh[skey]
        n = len(sh)
        n_cols = 10 + n

        # ── Title ──
        ws.append([f'单件装 {year_quarter}'])
        r_title = ws.max_row
        ws.row_dimensions[r_title].height = SLIP_TITLE_H
        _apply(ws.cell(r_title, 1), font=_f(SLIP_TITLE_SIZE), alignment=ALIGN_C)
        ws.merge_cells(start_row=r_title, start_column=1, end_row=r_title, end_column=n_cols)

        # ── Header ──
        ws.append(
            ['合同号：', '客款号', '厂款号：', '颜色', '客户色号', '英文颜色', '港口']
            + sh + ['总数', '工厂出厂期', '订单附加备注'])
        r_hdr = ws.max_row
        ws.row_dimensions[r_hdr].height = SLIP_HEADER_H
        _style_row(ws, r_hdr, n_cols, font=_f(SLIP_HEADER_SIZE), alignment=ALIGN_C,
                   border=BORDER_THIN)

        # ── Data lines (merged) ──
        r_first = ws.max_row + 1
        for ln in lines:
            size_vals = [ln['sizes'].get(s) or None for s in sh]
            ws.append(
                [ln['合同号'], ln['客款号'], ln['厂款号'], ln['颜色'],
                 ln['客户色号'], ln['英文颜色'], ln['港口']]
                + size_vals
                + [ln['总数'], ln['工厂出厂期'], None])
            r_data = ws.max_row
            ws.row_dimensions[r_data].height = SLIP_DATA_H
            _style_row(ws, r_data, n_cols, font=_f(SLIP_DATA_SIZE), alignment=ALIGN_C,
                       border=BORDER_THIN)
        r_last = ws.max_row

        # ── 合计 ──
        sum_row: list = ['合计'] + [None] * 6
        for ci in range(8, 8 + n):
            cl = get_column_letter(ci)
            sum_row.append(f'=SUM({cl}{r_first}:{cl}{r_last})')
        cl_total = get_column_letter(8 + n)
        sum_row.append(f'=SUM({cl_total}{r_first}:{cl_total}{r_last})')
        sum_row += [None, None]
        ws.append(sum_row)
        r_sum = ws.max_row
        ws.row_dimensions[r_sum].height = SLIP_DATA_H
        _style_row(ws, r_sum, n_cols, font=_f(SLIP_DATA_SIZE, bold=True),
                   alignment=ALIGN_C, fill=FILL_SUBTOTAL, border=BORDER_THIN)
        ws.merge_cells(start_row=r_sum, start_column=1, end_row=r_sum, end_column=7)

        # ── 注: 备注块 ──
        slip_ports = []
        for ln in lines:
            for p in ln['ports']:
                if p not in slip_ports:
                    slip_ports.append(p)
        note_text = SINGLE_SLIP_NOTE.format(port='/'.join(slip_ports), style=client_style)
        ws.append([note_text])
        r_notes = ws.max_row
        ws.row_dimensions[r_notes].height = SLIP_NOTES_H
        _apply(ws.cell(r_notes, 1), font=_f(SLIP_NOTES_SIZE), alignment=ALIGN_L)
        ws.merge_cells(start_row=r_notes, start_column=1, end_row=r_notes, end_column=n_cols)

        # ── Footer ──
        ws.append(['此单随货送回公司'])
        r_foot = ws.max_row
        ws.row_dimensions[r_foot].height = SLIP_FOOTER_H
        _apply(ws.cell(r_foot, 1), font=_f(SLIP_FOOTER_SIZE), alignment=ALIGN_C)
        ws.merge_cells(start_row=r_foot, start_column=1, end_row=r_foot, end_column=n_cols)

        # ── Blank rows ──
        ws.append([None])
        ws.append([None])
        block_end_rows.append(ws.max_row)

    # ── Page breaks + print setup (同 gen_special_port) ──
    for r in block_end_rows[:-1]:
        ws.row_breaks.append(Break(id=r))
    last_col = get_column_letter(10 + max_n)
    ws.print_area = f'A1:{last_col}{ws.max_row}'
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize   = 9
    ws.page_setup.fitToWidth  = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.3
    ws.page_margins.top  = ws.page_margins.bottom = 0.5


# ══════════════════════════════════════════════════════════════
# 辅助：格式 B → 格式 A（供 _write_source_sheet 使用）
# ══════════════════════════════════════════════════════════════

def _format_b_to_a(records_b):
    """Convert parse_port/parse_qty format-B records to _write_source_sheet format-A.

    Format B (parse_port/parse_qty output):
        '合同号', '厂款号', 'size_headers': [...], 'sizes': [...]

    Format A (_write_source_sheet input):
        '合同号：', '厂款号：', 'sizes_by_name': {size: qty}
    """
    result = []
    for rec in records_b:
        r = dict(rec)
        r['合同号：'] = rec.get('合同号', '')
        r['厂款号：'] = rec.get('厂款号', '')
        sizes_by_name = {}
        for name, val in zip(rec.get('size_headers', []), rec.get('sizes', [])):
            if val is not None and val != '':
                sizes_by_name[name] = val
        r['sizes_by_name'] = sizes_by_name
        result.append(r)
    return result


def process_from_waimao(waimao_path, output_shengchan, year_quarter=None,
                        pdf_output_dir=None):
    """Generate 生产/外发跟单 (6-tab) from an existing 外贸跟单 file.

    Reads the 港口明细汇总 and 数量明细汇总 sheets from the source file,
    then builds the 生产跟单 workbook (source sheets without actual dates +
    数量清单按颜色 + 3 special-port tabs).

    year_quarter: if provided, overrides the year/quarter auto-detected from the
                  source file's sheet name (e.g. '2026 Q2').
    pdf_output_dir: if set, generate 包装明细 PDFs (one per PO) into this folder.
    """
    wb_src = openpyxl.load_workbook(waimao_path, data_only=True)
    port_ws_src, qty_ws_src, auto_suffix, _ = detect_sources(wb_src)

    # Use user-provided year_quarter if given, otherwise auto-detect from source sheet name
    if year_quarter and str(year_quarter).strip():
        yq = str(year_quarter).strip()
    else:
        yq = auto_suffix.strip() or '未知'

    port_records_b = parse_port(port_ws_src)
    qty_records_b  = parse_qty(qty_ws_src)

    if not port_records_b or not qty_records_b:
        raise ValueError('外贸跟单文件中没有可生成的明细数据')

    # Build the 生产跟单 workbook
    wb_shengchan = Workbook()
    wb_shengchan.remove(wb_shengchan.active)

    # Write source sheets without actual dates
    port_ws_s = wb_shengchan.create_sheet(f'{SOURCE_PORT_PREFIX}{yq}')
    qty_ws_s  = wb_shengchan.create_sheet(f'{SOURCE_QTY_PREFIX}{yq}')
    _write_source_sheet(port_ws_s, _format_b_to_a(port_records_b),
                        include_port=True,  include_actual_dates=False)
    _write_source_sheet(qty_ws_s,  _format_b_to_a(qty_records_b),
                        include_port=False, include_actual_dates=False)

    # Parse back to format-B for the generator functions
    port_records = parse_port(port_ws_s)
    qty_records  = parse_qty(qty_ws_s)

    gen_qty_by_color(wb_shengchan, qty_records)
    gen_colour_class_summary(wb_shengchan, port_records)
    for tab_name, cfg in PORT_TAB_CONFIG.items():
        gen_special_port(wb_shengchan, port_records, tab_name, yq, cfg)
    gen_single_piece_slips(wb_shengchan, port_records, yq)

    wb_shengchan.save(output_shengchan)

    # ── 包装明细 PDFs (use format-A converted records) ──────────────
    if pdf_output_dir and str(pdf_output_dir).strip():
        gen_packing_pdfs(_format_b_to_a(port_records_b), str(pdf_output_dir).strip())


# ══════════════════════════════════════════════════════════════
# 包装明细 PDF 生成
# ══════════════════════════════════════════════════════════════

_PACKING_FOOTER = (
    '若排完柜车顺序后，港口有调整包装先后顺序，会另行通知包装部主管，'
    '以另行通知为准，大家要及时调整。'
)


def _qty_str(v):
    """Format a quantity value for display in the PDF."""
    if v is None or v == '':
        return ''
    try:
        f = float(v)
        return '' if f == 0 else (str(int(f)) if f == int(f) else str(f))
    except (TypeError, ValueError):
        return str(v)


def gen_packing_pdfs(port_records_a, output_dir):
    """Generate 包装明细 PDF files — one file per PO number.

    Accepts format-A port records (output of _records_from_po_csvs) where each
    record represents one (PO, 客款号, 客户色号, 港口) entry and carries
    ``sizes_by_name`` as a dict.

    Groups: PO → (客款号, 客户色号) → ports → one PDF page per colour group.
    """
    try:
        from reportlab.pdfgen import canvas as _rl_canvas
        from reportlab.lib.pagesizes import landscape as _landscape, A4 as _A4
        from reportlab.lib.units import mm as _mm
        from reportlab.pdfbase import pdfmetrics as _pdfmetrics
        from reportlab.pdfbase.cidfonts import UnicodeCIDFont as _UnicodeCIDFont
        from reportlab.lib import colors as _colors
    except ImportError:
        raise ImportError(
            '生成 PDF 需要 reportlab 库。\n请运行：pip install reportlab'
        )

    # Register an embedded TrueType CJK font so Chinese glyphs render everywhere.
    # Prefer SimHei → SimSun → MicrosoftYaHei → fallback to CID STSong-Light.
    from reportlab.pdfbase.ttfonts import TTFont as _TTFont
    import os as _os
    _TT_CANDIDATES = [
        ('PkgSimHei', r'C:\Windows\Fonts\simhei.ttf', 0),
        ('PkgSimSun', r'C:\Windows\Fonts\simsun.ttc', 0),
        ('PkgMSYH',   r'C:\Windows\Fonts\msyh.ttc',   0),
    ]
    FONT = None
    for _fname, _fpath, _idx in _TT_CANDIDATES:
        # Already registered in a previous call to this function?
        try:
            _pdfmetrics.getFont(_fname)
            FONT = _fname
            break
        except Exception:
            pass
        # Try to register from file
        if _os.path.exists(_fpath):
            try:
                _pdfmetrics.registerFont(_TTFont(_fname, _fpath, subfontIndex=_idx))
                FONT = _fname
                break
            except Exception:
                pass
    if FONT is None:
        # Last-resort CID font (no embedding — requires CJK font pack on viewer)
        try:
            _pdfmetrics.getFont('STSong-Light')
        except Exception:
            _pdfmetrics.registerFont(_UnicodeCIDFont('STSong-Light'))
        FONT = 'STSong-Light'

    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)

    # ── Group: po → (客款号, 客户色号) → [records] ─────────────────
    by_po = OrderedDict()
    for rec in port_records_a:
        po  = str(rec.get('合同号：', '') or rec.get('合同号', '') or '').strip()
        ck  = str(rec.get('客款号',   '') or '').strip()
        ccs = str(rec.get('客户色号', '') or '').strip()
        if not po:
            continue
        colour_key = (ck, ccs)
        by_po.setdefault(po, OrderedDict()).setdefault(colour_key, []).append(rec)

    PAGE_W, PAGE_H = _landscape(_A4)   # ≈841.89 × 595.28 pt
    mm = _mm
    LM = RM = 7 * mm
    TM = 8 * mm

    CW = PAGE_W - LM - RM   # content width ≈ 277mm in pt

    # Row heights (mm → pt via mm)
    R1H = 7  * mm   # date / pkg row
    R2H = 8  * mm   # 合同号 / 颜色 row
    R3H = 7  * mm   # column-header row
    DH  = 8  * mm   # data row
    SH  = 5  * mm   # spacer row height  (×3 per port block)
    TRH = 8  * mm   # 汇总 totals row
    FH  = 5  * mm   # footer text area

    # Fixed column widths
    SEQ_W  = 7  * mm   # left merged (blank – no carton data in CSV)
    CK_W   = 27 * mm   # 客款号 merged col
    CT_W   = 7  * mm   # carton# col (blank – not in CSV data)
    PORT_W = 20 * mm   # 港口 col
    DZ_W   = 6  * mm   # 单码装 col
    TOT_W  = 13 * mm   # 总数 col
    FIXED_W = SEQ_W + CK_W + CT_W + PORT_W + DZ_W + TOT_W

    for po, colour_groups in by_po.items():
        available_h = PAGE_H - TM - R1H - R2H - R3H - TRH - FH - 7 * mm
        max_ports = int(available_h // (DH + 6 * mm))
        pages = [(key, recs[start:start + max_ports], len(recs) > max_ports)
                 for key, recs in colour_groups.items()
                 for start in range(0, len(recs), max_ports)]
        total_pages = len(pages)
        customer_styles = list(dict.fromkeys(ck for ck, _ in colour_groups if ck))
        factory_styles = []
        for recs in colour_groups.values():
            for rec in recs:
                codes = _fac_pieces(rec) or [_rec_get(rec, '厂款号：'), rec.get('厂款号2')]
                for code in codes:
                    code = str(code or '').strip()
                    if code and code not in factory_styles:
                        factory_styles.append(code)
        name_parts = [po, '+'.join(customer_styles), '+'.join(factory_styles), '包装明细']
        pdf_name = re.sub(r'[\\/:*?"<>|]', '-', ' '.join(p for p in name_parts if p))
        pdf_path = output_dir / f'{pdf_name}.pdf'

        c = _rl_canvas.Canvas(str(pdf_path), pagesize=_landscape(_A4))
        c.setStrokeColor(_colors.black)
        c.setFillColor(_colors.black)

        for pg_idx, ((ck, ccs), recs, continued) in enumerate(pages):
            page_num    = pg_idx + 1
            factory_code = str(recs[0].get('厂款号：', '') or recs[0].get('厂款号', '') or '').strip()
            zh_color    = str(recs[0].get('颜色',    '') or '').strip()
            en_color    = str(recs[0].get('英文颜色', '') or '').strip()
            cust_full   = f'{ck}-{ccs}' if ccs else ck

            # Sorted union of sizes across all ports of this colour
            sh_all: list = []
            for r in recs:
                for s in r.get('sizes_by_name', {}):
                    if s not in sh_all:
                        sh_all.append(s)
            sh_all.sort(key=_size_sort_key)
            N = len(sh_all)

            # Size-column width: fill remaining space, capped at 18 mm
            avail_w  = CW - FIXED_W
            sz_w     = (avail_w / max(N, 1)) if N else avail_w
            sz_w     = min(sz_w, 18 * mm)
            leftover = avail_w - sz_w * N
            port_w   = PORT_W + leftover   # absorb spare width into 港口 column

            # Build column (x, width) list
            col_widths_list = [SEQ_W, CK_W, CT_W, port_w, DZ_W] + [sz_w] * N + [TOT_W]
            cols: list = []
            xc = LM
            for w in col_widths_list:
                cols.append((xc, w))
                xc += w

            I_SEQ = 0; I_CK = 1; I_CT = 2; I_PORT = 3; I_DZ = 4
            I_SZ0 = 5; I_TOT = 5 + N

            # Build port-row data (one entry per record = one port)
            port_rows = []
            for r in recs:
                szm = {k: v for k, v in (r.get('sizes_by_name') or {}).items()
                       if v not in (None, '', 0)}
                total = sum(v for v in szm.values() if isinstance(v, (int, float)))
                port_rows.append({
                    'port':     str(r.get('港口', '') or '').strip(),
                    'size_map': szm,
                    'total':    total,
                })

            n_ports    = len(port_rows)
            # Reserve the summary and footer before allocating handwriting space.
            SH = min(5 * mm, (available_h / n_ports - DH) / 3)
            data_sec_h = n_ports * (DH + 3 * SH)

            # ── Page drawing: y decreases downward ──────────────
            y = PAGE_H - TM   # current "top" y

            # ── Row 1: date / packaging row ──────────────────────
            rb = y - R1H
            c.setLineWidth(0.5)
            c.rect(LM, rb, CW, R1H, fill=0, stroke=1)
            c.setFont(FONT, 8)
            ty = rb + R1H * 0.38
            c.drawString(LM + 2, ty, '日期：')
            c.drawString(LM + CW * 0.30, ty, '包装负责人：')
            c.drawCentredString(LM + CW / 2, ty, f'{page_num} / {total_pages} 页')
            c.drawString(LM + CW * 0.72, ty, '装箱负责人：')
            y -= R1H

            # ── Row 2: 合同号 / 颜色 row ─────────────────────────
            rb = y - R2H
            c.setLineWidth(1.0)
            c.rect(LM, rb, CW, R2H, fill=0, stroke=1)
            d1 = LM + CW * 0.40
            d2 = d1 + 15 * mm
            d3 = d2 + CW * 0.20
            c.setLineWidth(1.0)
            for dx in (d1, d2, d3):
                c.line(dx, rb, dx, rb + R2H)
            ty = rb + R2H * 0.35
            c.setFont(FONT, 9)
            c.drawString(LM + 2, ty, f'合同号：{po}    厂款号：{factory_code}')
            c.drawCentredString((d1 + d2) / 2, ty, '颜色：')
            c.setFont(FONT, 10)
            c.drawCentredString((d2 + d3) / 2, ty, zh_color)
            c.drawCentredString((d3 + LM + CW) / 2, ty, en_color)
            y -= R2H

            # ── Row 3: column headers ─────────────────────────
            rb = y - R3H
            c.setLineWidth(0.5)
            for ci, (cx, cw) in enumerate(cols):
                c.rect(cx, rb, cw, R3H, fill=0, stroke=1)
            c.setFont(FONT, 7)
            c.drawCentredString(cols[I_PORT][0] + cols[I_PORT][1] / 2,
                                rb + R3H * 0.38, '港口：')
            for i, sz in enumerate(sh_all):
                cx, cw = cols[I_SZ0 + i]
                c.drawCentredString(cx + cw / 2, rb + R3H * 0.38, str(sz))
            cx_t, cw_t = cols[I_TOT]
            c.drawCentredString(cx_t + cw_t / 2, rb + R3H * 0.38, '总数')
            y -= R3H

            # ── Merged seq col (entire data section) ─────────
            cx0, cw0 = cols[I_SEQ]
            c.rect(cx0, y - data_sec_h, cw0, data_sec_h, fill=0, stroke=1)

            # ── Merged 客款号 col ─────────────────────────────
            cx1, cw1 = cols[I_CK]
            c.rect(cx1, y - data_sec_h, cw1, data_sec_h, fill=0, stroke=1)
            mid_y = y - data_sec_h / 2
            c.setFont(FONT, 7)
            c.drawCentredString(cx1 + cw1 / 2, mid_y + 3 * mm, '客款号：')
            c.drawCentredString(cx1 + cw1 / 2, mid_y - 1 * mm, cust_full)

            # ── Port rows ─────────────────────────────────────
            y_cur = y
            for pr in port_rows:
                db = y_cur - DH
                # Carton-number col (blank)
                cx_ct, cw_ct = cols[I_CT]
                c.rect(cx_ct, db, cw_ct, DH, fill=0, stroke=1)
                # Port name col
                cx_pt, cw_pt = cols[I_PORT]
                c.rect(cx_pt, db, cw_pt, DH, fill=0, stroke=1)
                c.setFont(FONT, 8)
                c.drawCentredString(cx_pt + cw_pt / 2, db + DH * 0.38, pr['port'])

                # Size qty cols
                for i, sz in enumerate(sh_all):
                    cx_s, cw_s = cols[I_SZ0 + i]
                    c.rect(cx_s, db, cw_s, DH, fill=0, stroke=1)
                    qs = _qty_str(pr['size_map'].get(sz, ''))
                    if qs:
                        c.setFont(FONT, 8)
                        c.drawCentredString(cx_s + cw_s / 2, db + DH * 0.38, qs)

                # Total col
                cx_t, cw_t = cols[I_TOT]
                c.rect(cx_t, db, cw_t, DH, fill=0, stroke=1)
                c.setFont(FONT, 8)
                c.drawCentredString(cx_t + cw_t / 2, db + DH * 0.38,
                                    _qty_str(pr['total']))

                # 单码装: merged over 1 data row + 3 spacer rows
                # (characters stacked vertically, upright)
                cx_dz, cw_dz = cols[I_DZ]
                dz_h = DH + 3 * SH
                c.rect(cx_dz, db - 3 * SH, cw_dz, dz_h, fill=0, stroke=1)
                c.setFont(FONT, 7)
                dz_chars = '单码装'
                dz_line_h = 9
                dz_cy = db - 3 * SH + dz_h / 2
                dz_y0 = dz_cy + dz_line_h * (len(dz_chars) - 1) / 2 - 2.5
                for dz_k, dz_ch in enumerate(dz_chars):
                    c.drawCentredString(cx_dz + cw_dz / 2,
                                        dz_y0 - dz_k * dz_line_h, dz_ch)

                # 3 spacer rows (carton + port + sizes + total; dz already merged)
                for sp in range(3):
                    sb = db - (sp + 1) * SH
                    c.rect(cx_ct, sb, cw_ct, SH, fill=0, stroke=1)
                    c.rect(cx_pt, sb, cw_pt, SH, fill=0, stroke=1)
                    for i in range(N):
                        cx_s, cw_s = cols[I_SZ0 + i]
                        c.rect(cx_s, sb, cw_s, SH, fill=0, stroke=1)
                    c.rect(cx_t, sb, cw_t, SH, fill=0, stroke=1)

                y_cur -= DH + 3 * SH

            # ── Totals (汇总) row ─────────────────────────────
            tr_b = y_cur - TRH
            # Merge seq+ck+ct+port+dz for '汇总' label
            merged_x = cols[I_SEQ][0]
            merged_w = sum(cols[i][1] for i in range(I_DZ + 1))
            c.setLineWidth(0.5)
            c.rect(merged_x, tr_b, merged_w, TRH, fill=0, stroke=1)
            c.setFont(FONT, 9)
            c.drawCentredString(merged_x + merged_w / 2, tr_b + TRH * 0.38,
                                '本页汇总' if continued else '汇总')

            for i, sz in enumerate(sh_all):
                cx_s, cw_s = cols[I_SZ0 + i]
                c.rect(cx_s, tr_b, cw_s, TRH, fill=0, stroke=1)
                sz_sum = sum(
                    (pr['size_map'].get(sz) or 0)
                    for pr in port_rows
                    if isinstance(pr['size_map'].get(sz), (int, float))
                )
                if sz_sum:
                    c.setFont(FONT, 8)
                    c.drawCentredString(cx_s + cw_s / 2, tr_b + TRH * 0.38,
                                        str(int(sz_sum)))

            cx_t, cw_t = cols[I_TOT]
            c.rect(cx_t, tr_b, cw_t, TRH, fill=0, stroke=1)
            grand_total = sum(pr['total'] for pr in port_rows
                              if isinstance(pr['total'], (int, float)))
            c.setFont(FONT, 8)
            c.drawCentredString(cx_t + cw_t / 2, tr_b + TRH * 0.38,
                                str(int(grand_total)))

            # ── Footer text ──────────────────────────────────
            c.setFont(FONT, 7)
            c.drawString(LM, tr_b - FH * 0.6, _PACKING_FOOTER)

            if pg_idx < total_pages - 1:
                c.showPage()

        c.save()


# ══════════════════════════════════════════════════════════════
# 主流程
# ══════════════════════════════════════════════════════════════

def check_missing_order_details(po_folder, order_detail_folder):
    """Return sorted PO numbers from the contract CSVs that have no 订单明细."""
    detail_pos = {str(key[0]).upper()
                  for key in load_order_detail_reference(order_detail_folder)}
    csv_pos = set()
    for path in latest_purchase_order_csvs(po_folder):
        for row in _read_po_csv(path):
            po = _clean(row.get('ORDER NUMBER'))
            if po:
                csv_pos.add(po.upper())
    return sorted(po for po in csv_pos if po not in detail_pos)


def process(po_folder, order_detail_folder, year_quarter, output_waimao, output_shengchan,
            pdf_output_dir=None, port_config_path=None):
    """Generate both output workbooks in one pass.

    Records are built once from the PO CSV folder + 订单明细 folder, then
    rendered into two separate files:

    output_waimao     — 外贸跟单: 4-sheet workbook
    output_shengchan  — 生产/外发跟单: 6-tab workbook
    pdf_output_dir    — if set, generate 包装明细 PDFs (one per PO) into this folder
    port_config_path  — optional path to CO港口分类.xlsx; drives port grouping
    """
    suffix = _clean(year_quarter)

    # ── Build records once ─────────────────────────────────────────
    csv_paths = latest_purchase_order_csvs(po_folder)
    field_lookup = _merge_field_lookups(
        load_order_detail_reference(order_detail_folder),
    )
    component_lookup = load_order_detail_components(order_detail_folder)
    port_records_csv, qty_records_csv = _records_from_po_csvs(
        csv_paths,
        field_lookup=field_lookup,
        component_lookup=component_lookup,
    )
    if not port_records_csv or not qty_records_csv:
        raise ValueError('采购订单 CSV 中没有可生成的明细数据')

    # ── 外贸跟单: 2-sheet source workbook with 实际出厂期/实际船期 ──
    wb_waimao = Workbook()
    wb_waimao.remove(wb_waimao.active)
    _write_source_sheet(
        wb_waimao.create_sheet(f'{SOURCE_PORT_PREFIX}{suffix}'),
        port_records_csv, include_port=True, include_actual_dates=True,
    )
    _write_source_sheet(
        wb_waimao.create_sheet(f'{SOURCE_QTY_PREFIX}{suffix}'),
        qty_records_csv, include_port=False, include_actual_dates=True,
    )
    # ── Load port config if provided ──────────────────────────────
    port_cfg_jiage = port_cfg_qita = None
    if port_config_path and Path(port_config_path).exists():
        try:
            port_cfg_jiage = load_port_config_jiage(port_config_path)
            port_cfg_qita  = load_port_config_qita(port_config_path)
        except Exception:
            pass  # fall back to legacy logic if config is unreadable

    gen_jiage_pai(wb_waimao, port_records_csv, port_cfg_jiage)
    gen_qita_fuliao(wb_waimao, port_records_csv, port_cfg_qita)
    wb_waimao.save(output_waimao)

    # ── 生产跟单: 6-tab workbook ───────────────────────────────────
    wb_shengchan = Workbook()
    wb_shengchan.remove(wb_shengchan.active)
    port_ws_s = wb_shengchan.create_sheet(f'{SOURCE_PORT_PREFIX}{suffix}')
    qty_ws_s  = wb_shengchan.create_sheet(f'{SOURCE_QTY_PREFIX}{suffix}')
    _write_source_sheet(port_ws_s, port_records_csv, include_port=True,  include_actual_dates=False)
    _write_source_sheet(qty_ws_s,  qty_records_csv,  include_port=False, include_actual_dates=False)

    # Parse the written source sheets into the format the generators expect
    port_records = parse_port(port_ws_s)
    qty_records  = parse_qty(qty_ws_s)

    gen_qty_by_color(wb_shengchan, qty_records)
    gen_colour_class_summary(wb_shengchan, port_records)
    for tab_name, cfg in PORT_TAB_CONFIG.items():
        gen_special_port(wb_shengchan, port_records, tab_name, year_quarter, cfg)
    gen_single_piece_slips(wb_shengchan, port_records, year_quarter)

    wb_shengchan.save(output_shengchan)

    # ── 包装明细 PDFs ────────────────────────────────────────────────
    if pdf_output_dir and str(pdf_output_dir).strip():
        gen_packing_pdfs(port_records_csv, str(pdf_output_dir).strip())


# ══════════════════════════════════════════════════════════════
# GUI
# ══════════════════════════════════════════════════════════════

def _default_output_paths(po_folder, year_quarter):
    """Return (waimao_path, shengchan_path) defaults based on the PO folder location."""
    parent = Path(po_folder).parent
    today  = datetime.now()
    tag    = f'{today.year}.{today.month}.{today.day}'
    yq     = _clean(year_quarter) or 'XXXX QX'
    waimao    = parent / f'CO {yq}  港口明细+数量明细汇总 {tag}.xlsx'
    shengchan = parent / f'外发跟单 CO {yq} {tag}.xlsx'
    return str(waimao), str(shengchan)


def build_into(parent):
    """Build the 跟单表生成工具 UI into an existing tkinter widget (Frame or Tk)."""
    saved = load_saved()
    p = {'padx': 10, 'pady': 6}

    # ── Row 0: 合同文件夹 (PO CSV folder) ─────────────────────────
    lbl_po = tk.Label(parent, text='合同文件夹：')
    lbl_po.grid(row=0, column=0, sticky='e', **p)
    po_var = tk.StringVar(value=saved.get('po_folder', ''))
    ent_po = tk.Entry(parent, textvariable=po_var, width=54)
    ent_po.grid(row=0, column=1, **p)

    waimao_var    = tk.StringVar(value=saved.get('output_waimao', ''))
    shengchan_var = tk.StringVar(value=saved.get('output_shengchan', ''))

    def _autofill_outputs():
        po = po_var.get().strip()
        yq = yq_var.get().strip()
        if po:
            w, s = _default_output_paths(po, yq)
            if not waimao_var.get().strip():
                waimao_var.set(w)
            if not shengchan_var.get().strip():
                shengchan_var.set(s)

    def browse_po():
        path = filedialog.askdirectory(title='选择合同文件夹（含 PO CSV）')
        if path:
            po_var.set(path)
            update_saved(po_folder=path)
            _autofill_outputs()

    btn_po = tk.Button(parent, text='文件夹…', command=browse_po)
    btn_po.grid(row=0, column=2, **p)

    # ── Row 1: 订单明细文件夹 ──────────────────────────────────────
    lbl_detail = tk.Label(parent, text='订单明细文件夹：')
    lbl_detail.grid(row=1, column=0, sticky='e', **p)
    detail_var = tk.StringVar(value=saved.get('order_detail_folder', ''))
    ent_detail = tk.Entry(parent, textvariable=detail_var, width=54)
    ent_detail.grid(row=1, column=1, **p)

    def browse_detail():
        path = filedialog.askdirectory(title='选择订单明细文件夹')
        if path:
            detail_var.set(path)
            update_saved(order_detail_folder=path)

    btn_detail = tk.Button(parent, text='文件夹…', command=browse_detail)
    btn_detail.grid(row=1, column=2, **p)

    # ── Row 2: 年度/季度 ───────────────────────────────────────────
    lbl_yq = tk.Label(parent, text='年度/季度：')
    lbl_yq.grid(row=2, column=0, sticky='e', **p)
    yq_var = tk.StringVar(value=saved.get('year_quarter', ''))
    ent_yq = tk.Entry(parent, textvariable=yq_var, width=20)
    ent_yq.grid(row=2, column=1, sticky='w', **p)
    tk.Label(parent, text='（例：2026 Q2）', fg='gray').grid(row=2, column=2, sticky='w', **p)

    # ── Row 3: 外贸跟单输出 ────────────────────────────────────────
    lbl_waimao = tk.Label(parent, text='外贸跟单输出：')
    lbl_waimao.grid(row=3, column=0, sticky='e', **p)
    ent_waimao = tk.Entry(parent, textvariable=waimao_var, width=54)
    ent_waimao.grid(row=3, column=1, **p)

    def browse_waimao():
        path = filedialog.asksaveasfilename(
            title='保存外贸跟单文件',
            defaultextension='.xlsx',
            filetypes=[('Excel 文件', '*.xlsx')]
        )
        if path:
            waimao_var.set(path)

    btn_waimao = tk.Button(parent, text='浏览…', command=browse_waimao)
    btn_waimao.grid(row=3, column=2, **p)

    # ── Row 4: 生产跟单输出 ────────────────────────────────────────
    tk.Label(parent, text='生产跟单输出：').grid(row=4, column=0, sticky='e', **p)
    tk.Entry(parent, textvariable=shengchan_var, width=54).grid(row=4, column=1, **p)

    def browse_shengchan():
        path = filedialog.asksaveasfilename(
            title='保存生产跟单文件',
            defaultextension='.xlsx',
            filetypes=[('Excel 文件', '*.xlsx')]
        )
        if path:
            shengchan_var.set(path)

    tk.Button(parent, text='浏览…', command=browse_shengchan).grid(row=4, column=2, **p)

    # ── Row 5: 包装明细 PDF 输出文件夹 ────────────────────────────────
    tk.Label(parent, text='包装明细PDF：').grid(row=5, column=0, sticky='e', **p)
    pdf_var = tk.StringVar(value=saved.get('pdf_output_dir', ''))
    ent_pdf = tk.Entry(parent, textvariable=pdf_var, width=54)
    ent_pdf.grid(row=5, column=1, **p)

    def browse_pdf_dir():
        path = filedialog.askdirectory(title='选择包装明细 PDF 输出文件夹')
        if path:
            pdf_var.set(path)
            update_saved(pdf_output_dir=path)

    tk.Button(parent, text='文件夹…', command=browse_pdf_dir).grid(row=5, column=2, **p)
    tk.Label(parent, text='（可选，留空则不生成 PDF）', fg='gray').grid(
        row=5, column=3, sticky='w', padx=(0, 10))

    # ── Row 6: 港口分类配置文件 ────────────────────────────────────
    tk.Label(parent, text='港口分类配置：').grid(row=6, column=0, sticky='e', **p)
    port_cfg_var = tk.StringVar(value=saved.get('port_config_path', ''))
    ent_cfg = tk.Entry(parent, textvariable=port_cfg_var, width=54)
    ent_cfg.grid(row=6, column=1, **p)

    def browse_port_cfg():
        path = filedialog.askopenfilename(
            title='选择港口分类配置文件',
            filetypes=[('Excel 文件', '*.xlsx'), ('所有文件', '*.*')],
        )
        if path:
            port_cfg_var.set(path)
            update_saved(port_config_path=path)

    tk.Button(parent, text='浏览…', command=browse_port_cfg).grid(row=6, column=2, **p)
    tk.Label(parent, text='（CO港口分类.xlsx，可选）', fg='gray').grid(
        row=6, column=3, sticky='w', padx=(0, 10))

    # ── Row 7: Separator + mode toggle ────────────────────────────
    ttk.Separator(parent, orient='horizontal').grid(
        row=7, column=0, columnspan=4, sticky='ew', padx=10, pady=(8, 2))

    use_waimao_src = tk.BooleanVar(value=False)
    waimao_src_var = tk.StringVar(value='')

    # Widgets that belong to normal (from-CSV) mode
    _csv_widgets = [lbl_po, ent_po, btn_po,
                    lbl_detail, ent_detail, btn_detail,
                    lbl_waimao, ent_waimao, btn_waimao]

    def _on_toggle(*_):
        if use_waimao_src.get():
            for w in _csv_widgets:
                w.config(state='disabled')
            lbl_src.grid()
            ent_src.grid()
            btn_src.grid()
        else:
            for w in _csv_widgets:
                w.config(state='normal')
            lbl_src.grid_remove()
            ent_src.grid_remove()
            btn_src.grid_remove()

    tk.Checkbutton(
        parent,
        text='使用已生成的外贸跟单文件直接生成生产跟单',
        variable=use_waimao_src,
        command=_on_toggle,
        font=('', 9),
    ).grid(row=8, column=0, columnspan=3, sticky='w', padx=10, pady=(2, 4))

    # ── Row 9: source file picker (hidden until checkbox checked) ──
    lbl_src = tk.Label(parent, text='外贸跟单源文件：')
    lbl_src.grid(row=9, column=0, sticky='e', **p)
    ent_src = tk.Entry(parent, textvariable=waimao_src_var, width=54)
    ent_src.grid(row=9, column=1, **p)

    def browse_waimao_src():
        path = filedialog.askopenfilename(
            title='选择已生成的外贸跟单文件',
            filetypes=[('Excel 文件', '*.xlsx'), ('所有文件', '*.*')]
        )
        if path:
            waimao_src_var.set(path)
            # Auto-fill shengchan output if still empty
            if not shengchan_var.get().strip():
                today = datetime.now()
                yq = yq_var.get().strip()
                tag = f'{today.year}.{today.month}.{today.day}'
                parent_dir = str(Path(path).parent)
                if yq:
                    fname = f'外发跟单 CO {yq} {tag}.xlsx'
                else:
                    fname = f'{Path(path).stem}_生产跟单_{today.month}.{today.day}.xlsx'
                shengchan_var.set(str(Path(parent_dir) / fname))

    btn_src = tk.Button(parent, text='浏览…', command=browse_waimao_src)
    btn_src.grid(row=9, column=2, **p)

    # Hide source-file row initially
    lbl_src.grid_remove()
    ent_src.grid_remove()
    btn_src.grid_remove()

    # ── Progress window ────────────────────────────────────────────
    def show_progress():
        top = parent.winfo_toplevel()
        win = tk.Toplevel(top)
        win.title('请稍候')
        win.resizable(False, False)
        win.transient(top)
        tk.Label(win, text='正在生成，请稍候…', font=('', 11)).pack(padx=40, pady=(22, 12))
        pb = ttk.Progressbar(win, mode='indeterminate', length=280)
        pb.pack(padx=40, pady=(0, 22))
        pb.start(12)
        win.protocol('WM_DELETE_WINDOW', lambda: None)
        win.grab_set()
        win.update_idletasks()
        x = top.winfo_x() + (top.winfo_width()  - win.winfo_width())  // 2
        y = top.winfo_y() + (top.winfo_height() - win.winfo_height()) // 2
        win.geometry(f'+{max(x, 0)}+{max(y, 0)}')
        return win

    # ── Generate ───────────────────────────────────────────────────
    def on_generate():
        out_s      = shengchan_var.get().strip()
        pdf_dir    = pdf_var.get().strip() or None
        cfg_path   = port_cfg_var.get().strip() or None

        if use_waimao_src.get():
            # ── Mode B: generate 生产跟单 from existing 外贸跟单 file ──
            src = waimao_src_var.get().strip()
            yq  = yq_var.get().strip()
            if not src:
                messagebox.showerror('错误', '请选择外贸跟单源文件'); return
            if not yq:
                messagebox.showerror('错误', '请输入年度/季度（例：2026 Q2）'); return
            if not out_s:
                messagebox.showerror('错误', '请指定生产跟单输出路径'); return

            gen_btn.config(state='disabled')
            progress = show_progress()
            result = {}

            def worker_b():
                try:
                    process_from_waimao(src, out_s, year_quarter=yq,
                                        pdf_output_dir=pdf_dir)
                    result['ok'] = True
                except PermissionError:
                    result['err'] = 'permission'
                except Exception:
                    result['err'] = traceback.format_exc()

            t = threading.Thread(target=worker_b, daemon=True)
            t.start()

            def poll_b():
                if t.is_alive():
                    parent.after(120, poll_b)
                    return
                progress.destroy()
                gen_btn.config(state='normal')
                if result.get('ok'):
                    update_saved(output_shengchan=out_s, year_quarter=yq)
                    msg = f'生产跟单（6 张）：\n{out_s}'
                    if pdf_dir:
                        msg += f'\n\n包装明细 PDF：\n{pdf_dir}'
                    messagebox.showinfo('完成', msg)
                elif result.get('err') == 'permission':
                    messagebox.showerror(
                        '无法保存',
                        '输出文件可能正在 Excel 中打开。\n请关闭后重新点击「生成」。')
                else:
                    messagebox.showerror('生成失败', result.get('err'))

            parent.after(120, poll_b)

        else:
            # ── Mode A: normal – generate both files from PO CSVs ──
            po     = po_var.get().strip()
            detail = detail_var.get().strip()
            yq     = yq_var.get().strip()

            if not po:
                messagebox.showerror('错误', '请先选择合同文件夹'); return
            if not yq:
                messagebox.showerror('错误', '请输入年度/季度（例：2026 Q2）'); return

            _autofill_outputs()
            out_w = waimao_var.get().strip()
            if not out_w or not out_s:
                messagebox.showerror('错误', '请指定输出文件路径'); return

            # ── Pre-check: every PO should have a 订单明细 ──────────
            try:
                missing = check_missing_order_details(po, detail or None)
            except Exception:
                missing = []
            if missing:
                if not messagebox.askyesno(
                        '缺少订单明细',
                        '以下 PO 未找到对应的订单明细文件，\n'
                        '相关字段（厂款号、颜色名称、备注等）可能不完整：\n\n'
                        + '\n'.join(missing)
                        + '\n\n是否仍要继续生成？'):
                    return

            gen_btn.config(state='disabled')
            progress = show_progress()
            result = {}

            def worker():
                try:
                    process(po, detail or None, yq, out_w, out_s,
                            pdf_output_dir=pdf_dir,
                            port_config_path=cfg_path)
                    result['ok'] = True
                except PermissionError:
                    result['err'] = 'permission'
                except Exception:
                    result['err'] = traceback.format_exc()

            t = threading.Thread(target=worker, daemon=True)
            t.start()

            def poll():
                if t.is_alive():
                    parent.after(120, poll)
                    return
                progress.destroy()
                gen_btn.config(state='normal')
                if result.get('ok'):
                    update_saved(
                        year_quarter=yq,
                        po_folder=po,
                        order_detail_folder=detail,
                        output_waimao=out_w,
                        output_shengchan=out_s,
                        pdf_output_dir=pdf_dir or '',
                        port_config_path=cfg_path or '',
                    )
                    msg = f'外贸跟单（4 张）：\n{out_w}\n\n生产跟单（6 张）：\n{out_s}'
                    if pdf_dir:
                        msg += f'\n\n包装明细 PDF：\n{pdf_dir}'
                    messagebox.showinfo('完成', msg)
                elif result.get('err') == 'permission':
                    messagebox.showerror(
                        '无法保存',
                        '输出文件可能正在 Excel 中打开。\n请关闭后重新点击「生成」。')
                else:
                    messagebox.showerror('生成失败', result.get('err'))

            parent.after(120, poll)

    gen_btn = tk.Button(
        parent, text='  生  成  ', command=on_generate,
        bg='#0078d4', fg='white', width=20, font=('', 11, 'bold'))
    gen_btn.grid(row=10, column=0, columnspan=3, pady=16)


def build_gui():
    """Standalone entry point — creates its own Tk root."""
    root = tk.Tk()
    root.title('CO港口/数量汇总工具')
    root.resizable(False, False)
    build_into(root)
    root.mainloop()


if __name__ == '__main__':
    build_gui()
