#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""装箱单生成工具 — 独立程序（确认无误后再合并入主程序）。

从 PO CSV + 订单明细 生成装箱资料表（PL），规则见
第三功能资料/装箱单分析与生成方案.md：
  - 每箱上限按客款号配置（UI 输入，自动记忆）
  - 标准模式：混码箱放颜色块开头，再按尺码整箱连跑
  - 小单：能装一箱就一箱
  - 箱高按装载率：>75% 大箱 / 40~75% 中箱 / <40% 小箱
  - ASOS/WHOLESALE 用 52×30 箱规，其余 52×32
  - 多港口 PO：一个文件多张表（1.0、2.0…），合同号加 -N 后缀
"""

import json
import math
import re
import sys
import threading
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk
from pathlib import Path
from collections import OrderedDict

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

import generate_tabs as gt
import pack_profiles as pp

if getattr(sys, 'frozen', False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).parent

CONFIG_FILE = APP_DIR / 'packing_config.json'

# ── 箱规 ──────────────────────────────────────────────────────
# 箱型尺寸由港口自动选定，无需用户输入：
#   ASOS EU / ASOS UK → 52×30  两种高度 (32 大 / 22 中·小)
#   其余所有港口      → 52×32  三种高度 (32 大 / 22 中 / 12 小)
DIMS_ASOS = {'大': (52, 30, 32), '中': (52, 30, 22), '小': (52, 30, 22)}
DIMS_STD  = {'大': (52, 32, 32), '中': (52, 32, 22), '小': (52, 32, 12)}
RATIO_MED   = 0.75   # >75% → 大箱
RATIO_SMALL = 0.40   # ≤40% → 小箱


BOX_SPEC_NOTE = '注意：纸箱尺寸选用请对应下方装箱单资料要求选用纸箱规格去装箱'
DESICCANT_NOTE = '注意：干燥剂放在防水袋里面，干燥剂用剩的要及时放回袋子并封存，避免受潮。'
WHOLESALE_NOTE = '1.此合同为WHOLESALE订单，不能短数或者多数，须按客人合同数量装箱'
ASOS_NOTE      = '2.注意：此合同为ASOS订单，箱规与其他订单不同'

# ── 字体（与原版装箱单一致）─────────────────────────────────
F_TNR22B  = Font(name='Times New Roman', size=22, bold=True)
F_TNR13B  = Font(name='Times New Roman', size=13, bold=True)
F_TNR13   = Font(name='Times New Roman', size=13)
F_TNR14B  = Font(name='Times New Roman', size=14, bold=True)
F_TNR11B  = Font(name='Times New Roman', size=11, bold=True)
F_TNR11BR = Font(name='Times New Roman', size=11, bold=True, color='FFFF0000')
F_SONG13B = Font(name='宋体', size=13, bold=True)
F_SONG14B = Font(name='宋体', size=14, bold=True)
F_SONG14BR = Font(name='宋体', size=14, bold=True, color='FFFF0000')
F_SONG15B = Font(name='宋体', size=15, bold=True)
F_SONG16B = Font(name='宋体', size=16, bold=True)
F_SONG16  = Font(name='宋体', size=16)
F_SONG18  = Font(name='宋体', size=18)
F_MING18  = Font(name='新細明體', size=18)
F_MING18C = Font(name='細明體', size=18)
F_MING16  = Font(name='細明體', size=16)
F_CAL16B  = Font(name='Calibri', size=16, bold=True)
F_SONG11  = Font(name='宋体', size=11)
F_SONG11B = Font(name='宋体', size=11, bold=True)
F_SONG12B = Font(name='宋体', size=12, bold=True)
F_TNR12B  = Font(name='Times New Roman', size=12, bold=True)
F_TNR12   = Font(name='Times New Roman', size=12)
F_TNR14   = Font(name='Times New Roman', size=14)
F_SONG14BR_Y = Font(name='宋体', size=14, bold=True, color='FFFF0000')
F_SONG13  = Font(name='宋体', size=13)

ALIGN_C   = Alignment(horizontal='center', vertical='center')
ALIGN_CW  = Alignment(horizontal='center', vertical='center', wrap_text=True)
ALIGN_L   = Alignment(horizontal='left',   vertical='center')
ALIGN_LW  = Alignment(horizontal='left',   vertical='center', wrap_text=True)
THIN        = Side(style='thin')
BORDER_ALL  = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BORDER_B    = Border(bottom=THIN)
BORDER_T    = Border(top=THIN)
BORDER_TB   = Border(top=THIN, bottom=THIN)
FILL_YELLOW = PatternFill('solid', start_color='FFFFFF00')
FILL_GOLD   = PatternFill('solid', start_color='FFFFE699')   # NW 常数底色
FILL_GREEN  = PatternFill('solid', start_color='FFE2EFDA')   # 右侧合计底色
CARTON_HEIGHT_FILLS = {12: PatternFill('solid', start_color='FFFCE4EC'), 22: FILL_GREEN}

# 数字格式（与原版一致）
NF_TEXT   = '@'
NF_INT    = '0_ '
NF_INT2   = '#,##0_);[Red](#,##0)'
NF_CBM    = '0.000_ '
NF_W      = '0.00_ '
NF_GROSS  = '0.00_);[Red]\\(0.00\\)'

# 固定列布局（与原版一致）：尺码区 F..X 共 19 列
COL_FROM, COL_SEP, COL_TO, COL_CNT = 2, 3, 4, 5
COL_SZ0   = 6           # F
N_SZ_SLOT = 19          # F..X
COL_PER   = 25          # Y 每箱数量
COL_TOT   = 26          # Z 总件数
COL_L, COL_W, COL_H = 27, 28, 29   # AA AB AC
COL_CBM   = 30          # AD
COL_NET   = 31          # AE
COL_TNET  = 32          # AF
COL_BOXW  = 33          # AG
COL_GROSS = 34          # AH
COL_TGROSS = 35         # AI


MAX_SIZES_PER_BOX = 4   # 每箱最多允许混装的不同尺码种数

# ══════════════════════════════════════════════════════════════
# 装箱算法
# ══════════════════════════════════════════════════════════════

def _pick_dims(per_box, cap_big, cap_mid, cap_small, dims):
    """Returns (kind, (L, W, H)) — picks smallest box that fits per_box."""
    if cap_small and per_box <= cap_small:
        return '小', dims['小']
    if cap_mid and per_box <= cap_mid:
        return '中', dims['中']
    if cap_mid or cap_small:   # at least one explicit threshold given
        return '大', dims['大']
    # Fallback: ratio-based when only cap_big is known
    ratio = per_box / cap_big if cap_big else 1
    if ratio > RATIO_MED:
        return '大', dims['大']
    if ratio > RATIO_SMALL:
        return '中', dims['中']
    return '小', dims['小']


def pack_colour(sizes_by_name, profile, cap_big=None):
    """Pack one colour by profile. Returns [{'sizes', 'per_box', 'cap'}].

    cap_big: user-specified pieces per large box; overrides profile tiers when given.
    余数混码箱放颜色块开头；尾箱 ≤ allowance 时并入前箱；小单一箱装完。
    """
    order = sorted(sizes_by_name, key=gt._size_sort_key)
    qty = {s: int(sizes_by_name[s]) for s in order if sizes_by_name[s]}
    order = [s for s in order if s in qty]
    if not qty:
        return []
    if cap_big:
        caps = {s: cap_big for s in order}
    else:
        caps = {s: pp.cap_for_size(profile, s, order) for s in order}
    allow = max(int(profile.allowance or 0), 0)
    total = sum(qty.values())
    min_cap = min(caps.values())
    # Small-order fast path only when sizes also fit within per-box limit
    if total <= min_cap + allow and len(order) <= MAX_SIZES_PER_BOX:
        return [{'sizes': dict(qty), 'per_box': total, 'cap': min_cap}]

    full = {s: qty[s] // caps[s] for s in order}
    rem  = {s: qty[s] % caps[s] for s in order}

    # 混码箱：余数按尺码顺序连续切；容量=箱内尺码的最小分级容量
    # 每箱最多 MAX_SIZES_PER_BOX 种尺码
    boxes = []
    cur, cur_n, cur_cap = {}, 0, None
    for s in order:
        r = rem[s]
        while r > 0:
            # Adding a new size would exceed the per-box size-variety limit → close box
            if s not in cur and len(cur) >= MAX_SIZES_PER_BOX and cur_n > 0:
                boxes.append({'sizes': cur, 'per_box': cur_n, 'cap': cur_cap})
                cur, cur_n, cur_cap = {}, 0, None
            eff_cap = caps[s] if cur_cap is None else min(cur_cap, caps[s])
            take = min(r, eff_cap - cur_n)
            if take <= 0:
                boxes.append({'sizes': cur, 'per_box': cur_n, 'cap': cur_cap})
                cur, cur_n, cur_cap = {}, 0, None
                continue
            cur[s] = cur.get(s, 0) + take
            cur_n += take
            cur_cap = eff_cap
            r -= take
            if cur_n >= eff_cap:
                boxes.append({'sizes': cur, 'per_box': cur_n, 'cap': cur_cap})
                cur, cur_n, cur_cap = {}, 0, None
    if cur_n:
        boxes.append({'sizes': cur, 'per_box': cur_n, 'cap': cur_cap or min_cap})

    # 尾箱并箱（allowance）— 仅当合并后不超过尺码种数限制
    if allow and len(boxes) >= 2 and boxes[-1]['per_box'] <= allow:
        prev = boxes[-2]
        merged_sizes = set(prev['sizes']) | set(boxes[-1]['sizes'])
        if (len(merged_sizes) <= MAX_SIZES_PER_BOX and
                prev['per_box'] + boxes[-1]['per_box'] <= prev['cap'] + allow):
            for s, v in boxes[-1]['sizes'].items():
                prev['sizes'][s] = prev['sizes'].get(s, 0) + v
            prev['per_box'] += boxes[-1]['per_box']
            boxes.pop()

    # 整箱按尺码顺序
    for s in order:
        for _ in range(full[s]):
            boxes.append({'sizes': {s: caps[s]}, 'per_box': caps[s], 'cap': caps[s]})
    return boxes


def merge_box_rows(boxes, start_no):
    """Merge consecutive identical boxes into rows.

    Returns (rows, next_no); row = {'from','to','n','sizes','per_box','cap'}.
    """
    rows = []
    no = start_no
    i = 0
    while i < len(boxes):
        j = i
        while j + 1 < len(boxes) and boxes[j + 1]['sizes'] == boxes[i]['sizes']:
            j += 1
        n = j - i + 1
        rows.append({
            'from': no, 'to': no + n - 1, 'n': n,
            'sizes': boxes[i]['sizes'], 'per_box': boxes[i]['per_box'],
            'cap': boxes[i].get('cap') or boxes[i]['per_box'],
        })
        no += n
        i = j + 1
    return rows, no


# ══════════════════════════════════════════════════════════════
# 数据准备
# ══════════════════════════════════════════════════════════════

def build_po_groups(po_folder, detail_folder):
    """Returns OrderedDict: po → OrderedDict(港口 → [colour records])."""
    csv_paths = gt.latest_purchase_order_csvs(po_folder)
    field_lookup = gt._merge_field_lookups(
        gt.load_order_detail_reference(detail_folder),
    )
    component_lookup = gt.load_order_detail_components(detail_folder)
    port_records, _ = gt._records_from_po_csvs(
        csv_paths, field_lookup=field_lookup, component_lookup=component_lookup)

    groups = OrderedDict()
    for rec in port_records:
        po = str(rec.get('合同号：', '') or '').strip()
        port = str(rec.get('港口', '') or '').strip()
        if not po:
            continue
        groups.setdefault(po, OrderedDict()).setdefault(port, []).append(rec)
    return groups


def _fac_display(rec, n_pieces=None):
    fac1 = str(rec.get('厂款号：') or '').strip()
    fac2 = str(rec.get('厂款号2') or '').strip()
    if fac2:
        return f'{fac1}+{fac2}（2件套）'
    if n_pieces and n_pieces > 1:
        return f'{fac1}（{n_pieces}件套）'
    return fac1


def _dest_display(recs):
    """Row-2 destination text, e.g. 'BASH 8094' / 'ASOS EU (WHOLESALE) - 8253'."""
    for rec in recs:
        code = str(rec.get('港口代码_detail') or '').strip()
        if code:
            return code
    raw = str(recs[0].get('destination_raw') or '').strip()
    if ' - ' in raw:
        code, name = raw.split(' - ', 1)
        return f'{name.strip()} {code.strip()}'
    return raw


def _is_wholesale(recs):
    return any('WHOLESALE' in str(r.get('订单附加备注') or '').upper() for r in recs)


def _is_asos(recs):
    return any('ASOS' in str(r.get('destination_raw') or '').upper() for r in recs)


def _dims_for_port(recs):
    """Return the correct dims_set dict based on destination (auto-detected)."""
    for r in recs:
        dest = str(r.get('destination_raw') or '').upper()
        if 'ASOS' in dest and ('EU' in dest or 'UK' in dest):
            return DIMS_ASOS
    return DIMS_STD


def _date_disp(recs):
    for rec in recs:
        iso = str(rec.get('工厂出厂期') or '').strip()
        m = re.match(r'\d{4}-(\d{1,2})-(\d{1,2})', iso)
        if m:
            return f'{int(m.group(1))}.{int(m.group(2))}'
    return ''


def _fmt_date(d):
    """Convert ISO 'YYYY-MM-DD' → 'M.D' shipping-date display; pass through anything else."""
    m = re.match(r'\d{4}-(\d{1,2})-(\d{1,2})', str(d or '').strip())
    if m:
        return f'{int(m.group(1))}.{int(m.group(2))}'
    return str(d or '').strip()


# ══════════════════════════════════════════════════════════════
_COL_A_CHARS = 20.53    # column A width in Excel character units
_EN_CHAR_WIDTH = 1.15   # TNR 14pt bold is wider than Calibri 11 normal


def _en_colour_row_height(text: str) -> float:
    """Row height (pt) needed to fully display `text` with wrap in column A."""
    if not text:
        return 20.0
    chars_per_line = max(1, int(_COL_A_CHARS / _EN_CHAR_WIDTH))
    n_lines = math.ceil(len(text) / chars_per_line)
    return max(20.0, n_lines * 20.0)


# 写一张装箱表
# ══════════════════════════════════════════════════════════════

def write_packing_sheet(ws, po_disp, recs, params, year_text, maker,
                        ship_date='', tags=None):
    """recs: colour-level records for one (PO, 港口). 版式复刻原版装箱单。

    params: {'profile': PackProfile, 'net': float|None,
             'w_big'/'w_mid'/'w_small': float|None}  — None 时用 profile 值。
    ship_date: 出货日期文本，写入 Z3:AI3。
    tags: {'吊牌': bool, '价钱牌': bool, '衣架': bool} — 勾选则写 √。
    """
    if tags is None:
        tags = {'吊牌': True, '价钱牌': True, '衣架': True}
    profile  = params['profile']
    net_pc   = params.get('net') or profile.net_k
    w_box    = {'大': params.get('w_big') or profile.gross_add,
                '中': params.get('w_mid') or profile.gross_add_mid,
                '小': params.get('w_small') or profile.gross_add_small}
    asos       = _is_asos(recs)
    dims_set   = _dims_for_port(recs)
    unit       = profile.unit

    # sheet-level size columns (union, sorted); 列位置固定 F..X
    all_sizes = []
    for rec in recs:
        for s in rec.get('sizes_by_name', {}):
            if s not in all_sizes:
                all_sizes.append(s)
    all_sizes.sort(key=gt._size_sort_key)
    n_sz = len(all_sizes)

    def cell(r, c, v=None, font=F_TNR13B, align=ALIGN_C, border=None,
             fill=None, numfmt=None):
        cl = ws.cell(r, c)
        if v is not None:
            cl.value = v
        cl.font = font
        cl.alignment = align
        if border is not None:
            cl.border = border
        if fill:
            cl.fill = fill
        if numfmt:
            cl.number_format = numfmt
        return cl

    def grid_row(r):
        """Apply full thin grid A..AI on a data-zone row."""
        for c in range(1, COL_TGROSS + 1):
            ws.cell(r, c).border = BORDER_ALL

    # ── 表头区（行1-11，与原版坐标一致）───────────────────────
    ws.merge_cells('A1:AH1')
    cell(1, 1, '装箱资料表', font=F_TNR22B)
    ws.row_dimensions[1].height = 27.0

    dest_txt = _dest_display(recs)
    cell(2, 7, dest_txt, font=F_TNR13B, align=ALIGN_L)        # G2
    cell(2, 31, year_text, font=F_TNR13B)                     # AE2
    ws.row_dimensions[2].height = 25.5

    note_parts = []
    if _is_wholesale(recs):
        note_parts.append(WHOLESALE_NOTE)
        if asos:
            note_parts.append(ASOS_NOTE)
    if note_parts:
        ws.merge_cells(start_row=2, start_column=11, end_row=3, end_column=22)  # K2:V3
        cell(2, 11, '\n'.join(note_parts), font=F_SONG14BR_Y,
             align=ALIGN_LW, fill=FILL_YELLOW)

    cell(3, 1, '合同号:', font=F_TNR13B, align=ALIGN_L)
    ws.merge_cells(start_row=3, start_column=4, end_row=3, end_column=6)
    cell(3, 4, po_disp, font=F_TNR13B, border=BORDER_B, numfmt=NF_TEXT)
    cell(3, 7, str(recs[0].get('港口') or ''), font=F_SONG13B, align=ALIGN_L)
    ws.merge_cells(start_row=3, start_column=26, end_row=3, end_column=35)  # Z3:AI3 — 出货日期
    cell(3, 26, _fmt_date(ship_date) if ship_date else None, align=ALIGN_L, border=BORDER_ALL)
    ws.row_dimensions[3].height = 17.25
    ws.row_dimensions[4].height = 12.95

    cell(5, 1, '款  号:', font=F_TNR13B, align=ALIGN_L)
    ws.merge_cells(start_row=5, start_column=4, end_row=5, end_column=6)
    cell(5, 4, str(recs[0].get('客款号') or ''), font=F_TNR13B,
         border=BORDER_B, numfmt=NF_TEXT)

    # 外箱尺寸 / 干燥剂用量（O5:S5 / T5:W5，行6-8 同列）
    ws.merge_cells('O5:S5'); ws.merge_cells('T5:W5')
    cell(5, 15, '外箱尺寸', font=F_SONG15B, border=BORDER_ALL)
    cell(5, 20, '干燥剂用量', font=F_SONG15B, border=BORDER_ALL)
    for i, kind in enumerate(('大', '中', '小')):
        L, W, H = dims_set[kind]
        rr = 6 + i
        ws.merge_cells(start_row=rr, start_column=15, end_row=rr, end_column=19)
        ws.merge_cells(start_row=rr, start_column=20, end_row=rr, end_column=23)
        cell(rr, 15, f'{L}x{W}x{H}（{kind}规格尺寸）', font=F_TNR14B, border=BORDER_ALL)
        g = '5g' if kind == '小' else '10g'
        cell(rr, 20, f'{g}（1包）每箱', font=F_SONG14B, border=BORDER_ALL)
    ws.merge_cells('Y6:AH10')
    _bd, _md, _sd = dims_set['大'], dims_set['中'], dims_set['小']
    box_spec_text = (
        f'外箱尺寸：{_bd[0]}x{_bd[1]}x{_bd[2]} （大规格尺寸）\n'
        f'外箱尺寸：{_md[0]}x{_md[1]}x{_md[2]} （中规格尺寸）\n'
        f'外箱尺寸：{_sd[0]}x{_sd[1]}x{_sd[2]} （小规格尺寸）\n'
        + BOX_SPEC_NOTE
    )
    cell(6, 25, box_spec_text, font=F_SONG14B, align=ALIGN_LW, border=BORDER_ALL)
    ws.row_dimensions[5].height = 16.5
    ws.row_dimensions[6].height = 18.0
    ws.row_dimensions[7].height = 18.0
    ws.row_dimensions[8].height = 24.0

    cell(7, 1, '包装方法:', font=F_TNR13B, align=ALIGN_L)

    # 吊牌 / 价钱牌 / 衣架（行9 — 勾选由 UI 控制）
    cell(9, 1, '吊牌', font=F_TNR13B, align=ALIGN_L)
    cell(9, 2, '√' if tags.get('吊牌') else None, font=F_TNR13B, border=BORDER_ALL)
    ws.merge_cells('F9:G9')
    cell(9, 6, '价钱牌', font=F_TNR13B)
    cell(9, 8, '√' if tags.get('价钱牌') else None, font=F_TNR13B, border=BORDER_ALL)
    cell(9, 10, '衣架', font=F_TNR13B,
         align=Alignment(horizontal='right', vertical='center'))
    cell(9, 11, '√' if tags.get('衣架') else None, font=F_TNR13B, border=BORDER_ALL)
    ws.merge_cells('O9:W10')
    cell(9, 15, DESICCANT_NOTE, font=F_SONG14BR, align=ALIGN_CW,
         border=BORDER_ALL, fill=FILL_YELLOW)
    # 右上常数区：单件净重（尺寸已在左侧箱规说明区显示，不重复）
    cell(7, 35, 'NW:', font=F_TNR14)
    cell(7, 36, net_pc, font=F_SONG14B, fill=FILL_GOLD, numfmt=NF_CBM)
    ws.row_dimensions[9].height = 24.0
    ws.row_dimensions[10].height = 18.75
    ws.row_dimensions[11].height = 15.0

    # ── 明细表头（行12-13 双行合并，与原版一致）───────────────
    HDR_ROW = 12
    ws.merge_cells(start_row=HDR_ROW, start_column=COL_FROM,
                   end_row=HDR_ROW + 1, end_column=COL_TO)         # B12:D13
    for c in [1, COL_CNT] + list(range(COL_SZ0, COL_SZ0 + N_SZ_SLOT)) + \
             list(range(COL_PER, COL_TGROSS + 1)):
        ws.merge_cells(start_row=HDR_ROW, start_column=c,
                       end_row=HDR_ROW + 1, end_column=c)
    cell(HDR_ROW, 1, '颜色', font=F_TNR13B, border=BORDER_ALL)
    cell(HDR_ROW, COL_FROM, '箱号', font=F_TNR13B, border=BORDER_ALL)
    cell(HDR_ROW, COL_CNT, '总箱\n数', font=F_SONG13B, align=ALIGN_CW, border=BORDER_ALL)
    for i in range(N_SZ_SLOT):
        v = all_sizes[i] if i < n_sz else None
        cell(HDR_ROW, COL_SZ0 + i, v, font=F_TNR14B, border=BORDER_ALL)
    cell(HDR_ROW, COL_PER, '每箱数量', font=F_TNR11B, align=ALIGN_CW, border=BORDER_ALL)
    cell(HDR_ROW, COL_TOT, '总件数', font=F_TNR13B, align=ALIGN_CW, border=BORDER_ALL,
         numfmt=NF_W)
    cell(HDR_ROW, COL_L, '长', font=F_SONG13B, align=ALIGN_CW, border=BORDER_ALL)
    cell(HDR_ROW, COL_W, '宽', font=F_SONG13B, align=ALIGN_CW, border=BORDER_ALL)
    cell(HDR_ROW, COL_H, '高', font=F_SONG13B, align=ALIGN_CW, border=BORDER_ALL)
    cell(HDR_ROW, COL_CBM, 'CBM', font=F_SONG13B, align=ALIGN_CW, border=BORDER_ALL,
         numfmt=NF_CBM)
    cell(HDR_ROW, COL_NET, '净重', font=F_TNR11B, align=ALIGN_CW, border=BORDER_ALL,
         numfmt=NF_W)
    cell(HDR_ROW, COL_TNET, '总净重', font=F_TNR11B, align=ALIGN_CW, border=BORDER_ALL,
         numfmt=NF_W)
    cell(HDR_ROW, COL_BOXW, '纸箱重量', font=F_SONG11B, align=ALIGN_CW, border=BORDER_ALL,
         numfmt=NF_W)
    cell(HDR_ROW, COL_GROSS, '毛重', font=F_TNR11B, align=ALIGN_CW, border=BORDER_ALL,
         numfmt=NF_W)
    cell(HDR_ROW, COL_TGROSS, '总毛重', font=F_TNR11B, align=ALIGN_CW, border=BORDER_ALL,
         numfmt=NF_W)
    # 双行表头补边框
    for c in range(1, COL_TGROSS + 1):
        ws.cell(HDR_ROW, c).border = BORDER_ALL
        ws.cell(HDR_ROW + 1, c).border = BORDER_ALL
    ws.row_dimensions[HDR_ROW].height = 15.0
    ws.row_dimensions[HDR_ROW + 1].height = 18.0

    # ── 装箱明细（行14起）────────────────────────────────────
    sz0_l   = get_column_letter(COL_SZ0)                 # F
    szX_l   = get_column_letter(COL_SZ0 + N_SZ_SLOT - 1)  # X
    per_l   = get_column_letter(COL_PER)                 # Y
    cnt_l   = get_column_letter(COL_CNT)                 # E
    tot_l   = get_column_letter(COL_TOT)                 # Z
    net_l   = get_column_letter(COL_NET)                 # AE
    tnet_l  = get_column_letter(COL_TNET)                # AF
    boxw_l  = get_column_letter(COL_BOXW)                # AG
    gross_l  = get_column_letter(COL_GROSS)              # AH
    tgross_l = get_column_letter(COL_TGROSS)             # AI

    r = HDR_ROW + 2
    box_no = 1
    first_data_row = r
    colour_row_map = []   # (rec, [(row, n_boxes), ...]) for summary formulas
    n_blocks = 0
    for rec in recs:
        boxes = pack_colour(rec.get('sizes_by_name', {}), profile,
                            cap_big=params.get('cap_big'))
        if not boxes:
            continue
        if n_blocks:
            grid_row(r)          # 颜色块之间空一行（保留网格线）
            r += 1
        n_blocks += 1
        rows, box_no = merge_box_rows(boxes, box_no)
        colour_cn = str(rec.get('颜色') or '').strip()
        colour_en = str(rec.get('英文颜色') or '').strip().upper()
        row_list = []
        block_fr = r
        for k, row in enumerate(rows):
            grid_row(r)
            if k == 0:
                cell(r, 1, colour_cn, font=F_SONG14B, border=BORDER_ALL,
                     fill=FILL_YELLOW)
            elif k == 1:
                cell(r, 1, colour_en, font=F_TNR14B, align=ALIGN_CW,
                     border=BORDER_ALL, fill=FILL_YELLOW)
            cell(r, COL_FROM, str(row['from']), font=F_TNR13B,
                 border=BORDER_ALL, numfmt=NF_TEXT)
            cell(r, COL_TO, str(row['to']), font=F_TNR13B,
                 border=BORDER_ALL, numfmt=NF_TEXT)
            cell(r, COL_CNT, f'=D{r}-B{r}+1', font=F_TNR13B,
                 border=BORDER_ALL, numfmt=NF_INT)
            for i, s in enumerate(all_sizes):
                v = row['sizes'].get(s)
                if v:
                    cell(r, COL_SZ0 + i, v, font=F_TNR13B, border=BORDER_ALL)
            per = row['per_box']
            kind, (L, W, H) = _pick_dims(per, params.get('cap_big'),
                                          params.get('cap_mid'), params.get('cap_small'),
                                          dims_set)
            cell(r, COL_PER, f'=SUM({sz0_l}{r}:{szX_l}{r})', font=F_TNR13,
                 border=BORDER_ALL)
            cell(r, COL_TOT, f'={per_l}{r}*{cnt_l}{r}', font=F_TNR13,
                 border=BORDER_ALL, numfmt=NF_INT)
            carton_fill = CARTON_HEIGHT_FILLS.get(H)
            for c, dv in ((COL_L, L), (COL_W, W), (COL_H, H)):
                cell(r, c, dv, font=F_TNR13, border=BORDER_ALL, fill=carton_fill, numfmt=NF_INT2)
            cell(r, COL_CBM,
                 f'={get_column_letter(COL_H)}{r}*{get_column_letter(COL_W)}{r}'
                 f'*{get_column_letter(COL_L)}{r}/1000000*{cnt_l}{r}',
                 font=F_TNR13, border=BORDER_ALL, numfmt=NF_CBM)
            cell(r, COL_NET, f'={net_pc}*{per_l}{r}', font=F_TNR11B,
                 border=BORDER_ALL, numfmt=NF_W)
            cell(r, COL_TNET, f'={net_l}{r}*{cnt_l}{r}', font=F_TNR11B,
                 border=BORDER_ALL, numfmt=NF_W)
            cell(r, COL_BOXW, w_box[kind], font=F_TNR11BR,
                 border=BORDER_ALL, fill=carton_fill, numfmt=NF_W)
            cell(r, COL_GROSS, f'={net_l}{r}+{boxw_l}{r}', font=F_TNR11B,
                 border=BORDER_ALL, numfmt=NF_W)
            cell(r, COL_TGROSS, f'={gross_l}{r}*{cnt_l}{r}', font=F_TNR11B,
                 border=BORDER_ALL, numfmt=NF_GROSS)
            ws.row_dimensions[r].height = (
                _en_colour_row_height(colour_en) if k == 1
                else (18.0 if r % 2 == 0 else 19.0)
            )
            row_list.append((r, row['n'], row['sizes']))
            r += 1
        # 颜色块只有 1 行时补一行放英文色名（颜色列必须中英都有，无英文色名则留空行）
        if len(rows) == 1:
            grid_row(r)
            cell(r, 1, colour_en, font=F_TNR14B, align=ALIGN_CW,
                 border=BORDER_ALL, fill=FILL_YELLOW)
            ws.row_dimensions[r].height = _en_colour_row_height(colour_en)
            r += 1
        # 英文色名跨块内第2-3行合并（与原版 A15:A16 一致）
        if len(rows) >= 3:
            ws.merge_cells(start_row=block_fr + 1, start_column=1,
                           end_row=block_fr + 2, end_column=1)
        # 补空行：右侧 helper 需要至少 3 行（色名 + 标题 + 数值）
        while r - block_fr < 3:
            grid_row(r)
            ws.row_dimensions[r].height = 18.0 if r % 2 == 0 else 19.0
            r += 1
        # 右侧小计区（总净重/总毛重/总件数/箱数/总CBM），与原版 AJ..AN 一致
        block_lr = r - 1
        n_phys = block_lr - block_fr + 1
        helper_cols = (
            (36, '总净重', tnet_l,   NF_W),
            (37, '总毛重', tgross_l, NF_W),
            (38, '总件数', tot_l,    NF_INT),
            (39, '箱数',   cnt_l,    NF_INT),
            (40, '总CBM',  get_column_letter(COL_CBM), NF_W),
        )
        # 色名行：block_lr-2；标题行（单行）：block_lr-1；数值行：block_lr
        if n_phys >= 3:
            _cn_r = block_lr - 2
            ws.merge_cells(start_row=_cn_r, start_column=36,
                           end_row=_cn_r, end_column=40)
            cell(_cn_r, 36, colour_cn, font=F_SONG14B,
                 border=BORDER_ALL, fill=FILL_YELLOW)
        if n_phys >= 2:
            lab_r = block_lr - 1
            for c, lbl, _src, nf in helper_cols:
                cell(lab_r, c, lbl, font=F_SONG11, border=BORDER_ALL, numfmt=nf)
        for ci, (c, _lbl, src, nf) in enumerate(helper_cols):
            cell(block_lr, c, f'=SUM({src}{block_fr}:{src}{block_lr})',
                 font=F_TNR12B if ci == 0 else F_SONG12B,
                 border=BORDER_ALL, fill=FILL_GREEN, numfmt=nf)
        colour_row_map.append((rec, row_list))
    last_data_row = r - 1

    # ── 合计区（紧跟数据行，与原版一致）───────────────────────
    fr, lr = first_data_row, last_data_row
    sum_row = r
    cell(r, 4, '共', font=F_TNR13B)
    cell(r, 5, f'=SUM({cnt_l}{fr}:{cnt_l}{lr})', font=F_TNR13B, numfmt=NF_INT2)
    cell(r, 6, '箱', font=F_TNR13B)
    cell(r, COL_PER, '共', font=F_TNR13B)
    cell(r, COL_TOT, f'=SUM({tot_l}{fr}:{tot_l}{lr})', font=F_TNR13B, numfmt=NF_INT)
    cell(r, COL_TNET, f'=SUM({tnet_l}{fr}:{tnet_l}{lr})', font=F_TNR13B, numfmt=NF_W)
    cell(r, COL_TGROSS, f'=SUM({tgross_l}{fr}:{tgross_l}{lr})', font=F_TNR13B, numfmt=NF_W)
    # 右侧小计区总计（SUM 范围内只有各色块小计单元格，文本标签被忽略）
    for c, nf in ((36, NF_W), (37, NF_W), (38, NF_INT), (39, NF_INT), (40, NF_W)):
        cl = get_column_letter(c)
        cell(r, c, f'=SUM({cl}{fr}:{cl}{lr})', font=F_TNR12, numfmt=nf)
    ws.row_dimensions[r].height = 18.0
    r += 1
    def _fh_value(rr, value, border):
        ws.merge_cells(start_row=rr, start_column=6, end_row=rr, end_column=8)
        cell(rr, 6, value, font=F_TNR13B, border=border, numfmt=NF_INT2)
        ws.cell(rr, 7).border = border
        ws.cell(rr, 8).border = border

    cell(r, 1, '制单总数:', font=F_TNR13B, align=ALIGN_L)
    _fh_value(r, f'={tot_l}{sum_row}', BORDER_B)
    cell(r, 9, unit, font=F_TNR13B, align=ALIGN_L)
    cell(r, COL_TOT, '生产跟单:', font=F_TNR13B, align=ALIGN_L)
    r += 1
    cell(r, 1, '总箱数:', font=F_TNR13B, align=ALIGN_L)
    _fh_value(r, f'={cnt_l}{sum_row}', BORDER_TB)
    cell(r, 9, '箱', font=F_TNR13B, align=ALIGN_L)
    cell(r, COL_TOT, '船  务:', font=F_TNR13B, align=ALIGN_L)
    r += 1
    cell(r, 1, '制表人:', font=F_TNR13B, align=ALIGN_L)
    _fh_value(r, maker, BORDER_TB)
    cell(r, COL_TOT, '经  理:', font=F_TNR13B, align=ALIGN_L)
    r += 2

    # ── 订单数量汇总（两份）────────────────────────────────────
    # 第1份：来自装箱单（可对照核查）；第2份：来自订单原数（用于人工核查）
    raw = str(recs[0].get('destination_raw') or '').strip()
    port_cn = str(recs[0].get('港口') or '')
    if ' - ' in raw:
        city_en = raw.split(' - ', 1)[1].strip()
        for kw in ('DISTRIBUTION', 'CENTRE', ' - ', ' ('):
            idx = city_en.upper().find(kw.upper())
            if idx > 0:
                city_en = city_en[:idx].strip()
    else:
        city_en = ''
    sum_title = f'{city_en} {port_cn}'.strip()
    for _copy in range(2):
        use_src = (_copy == 1)   # True → 订单原数；False → 装箱单数量
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=COL_TNET)
        cell(r, 1, sum_title, font=F_CAL16B, border=BORDER_T)
        ws.row_dimensions[r].height = 18.0
        r += 1
        cell(r, 1, '主布颜色', font=F_SONG13B, border=BORDER_ALL)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
        cell(r, 2, '客色号', font=F_SONG13B, border=BORDER_ALL)
        cell(r, 5, '色号', font=F_SONG12B, border=BORDER_ALL)
        for i in range(N_SZ_SLOT):
            v = all_sizes[i] if i < n_sz else None
            cell(r, COL_SZ0 + i, v, font=F_TNR12B, border=BORDER_ALL)
        ws.row_dimensions[r].height = 22.0
        hdr_r = r
        r += 1
        for rec, row_list in colour_row_map:
            cell(r, 1, str(rec.get('颜色') or ''), font=F_SONG13B,
                 align=ALIGN_CW, border=BORDER_ALL)
            ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=4)
            cell(r, 2, str(rec.get('英文颜色') or '').upper(), font=F_SONG13,
                 align=Alignment(vertical='center', wrap_text=True),
                 border=BORDER_ALL, numfmt=NF_TEXT)
            _raw_cc = str(rec.get('客户色号') or '').strip()
            _m = re.search(r'(\d+)$', _raw_cc)
            cell(r, 5, _m.group(1) if _m else _raw_cc, font=F_SONG12B,
                 align=Alignment(vertical='center'), border=BORDER_ALL,
                 numfmt=NF_TEXT)
            if use_src:
                src_sizes = rec.get('sizes_by_name', {})
                for i, s in enumerate(all_sizes):
                    v = src_sizes.get(s)
                    qty = int(v) if v else None
                    cell(r, COL_SZ0 + i, qty, font=F_SONG13B, border=BORDER_ALL)
            else:
                for i, s in enumerate(all_sizes):
                    terms = []
                    for rr, n, szs in row_list:
                        if szs.get(s):
                            col = get_column_letter(COL_SZ0 + i)
                            terms.append(f'{col}{rr}*E{rr}' if n > 1 else f'{col}{rr}')
                    cell(r, COL_SZ0 + i, ('=' + '+'.join(terms)) if terms else None,
                         font=F_SONG13B, border=BORDER_ALL)
            for i in range(n_sz, N_SZ_SLOT):
                ws.cell(r, COL_SZ0 + i).border = BORDER_ALL
            cell(r, COL_TOT, '=', font=F_SONG13B,
                 align=Alignment(horizontal='right', vertical='center'), numfmt=NF_W)
            cell(r, COL_TNET, f'=SUM({sz0_l}{r}:{per_l}{r})', font=F_SONG13B)
            cell(r, COL_TGROSS, unit, font=F_SONG13B, align=ALIGN_L)
            ws.row_dimensions[r].height = 22.0
            r += 1
        r += 1
        cell(r, COL_TNET,
             f'=SUM({tnet_l}{hdr_r + 1}:{tnet_l}{r - 2})',
             font=F_TNR14B, numfmt=NF_INT)
        r += 1
        if _copy == 0:
            print_last_row = r - 1   # 打印区至第1份汇总末行，不含第2份原始数据

    # ── 列宽（与原版一致）/ 打印 ──────────────────────────────
    widths = {
        'A': 20.53, 'B': 4.88, 'C': 1.26, 'D': 4.88, 'E': 5.73,
        'F': 7.78, 'I': 7.78, 'L': 7.78, 'P': 6.18, 'R': 6.18,
        'Z': 7.26, 'AA': 4.63, 'AD': 6.63, 'AE': 6.26, 'AF': 8.88,
        'AG': 5.5, 'AH': 6.76, 'AI': 8.76,
        'AJ': 10.47, 'AK': 10.06, 'AL': 9.26, 'AM': 6.88, 'AN': 8.76,
    }
    for letter, w in widths.items():
        ws.column_dimensions[letter].width = w
    # 打印设置：A4 横向，列适合一页宽，行不限页数，近零页边距，行1-13重复
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = 9          # A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1         # 所有列压缩到一页宽
    ws.page_setup.fitToHeight = 0        # 行数不限（按内容自动续页）
    ws.page_margins.left = 0.0785
    ws.page_margins.right = 0.157
    ws.page_margins.top = 0.0
    ws.page_margins.bottom = 0.0
    ws.page_margins.header = 0.0
    ws.page_margins.footer = 0.0
    ws.print_area = f'A1:AN{print_last_row}'
    ws.print_title_rows = '$1:$13'
    ws.sheet_view.zoomScale = 70
    _solidify_merged_borders(ws)


def _solidify_merged_borders(ws):
    """openpyxl 保存时会丢弃合并区覆盖单元格的样式，导致合并框边线缺失。
    把覆盖单元格替换为带上左单元格边框的真实 Cell（与 Excel 自己保存的
    文件一致），使合并区四边完整显示。"""
    from copy import copy as _copy
    from openpyxl.cell.cell import Cell, MergedCell
    for rng in ws.merged_cells.ranges:
        tl = ws.cell(rng.min_row, rng.min_col)
        bd = _copy(tl.border)
        for rr in range(rng.min_row, rng.max_row + 1):
            for cc in range(rng.min_col, rng.max_col + 1):
                if rr == rng.min_row and cc == rng.min_col:
                    continue
                if isinstance(ws.cell(rr, cc), MergedCell):
                    real = Cell(ws, row=rr, column=cc)
                    real.border = bd
                    ws._cells[(rr, cc)] = real


def sanitize_filename(name):
    return re.sub(r'[\\/:*?"<>|]', '-', name).strip()


def generate_packing_lists(po_folder, detail_folder, out_dir, year_quarter,
                           maker, params_by_style, ship_dates=None,
                           comp_pack=None):
    """Generate one workbook per PO. Returns (n_files, [warnings]).
    ship_dates: {'PO|港口': date_str} per-sheet override.
    comp_pack: {客款号: {'comp': str, 'pack': str}} written to hidden _meta sheet.
    """
    groups = build_po_groups(po_folder, detail_folder)
    if not groups:
        raise ValueError('采购订单 CSV 中没有可生成的数据')
    out = Path(out_dir)
    out.mkdir(parents=True, exist_ok=True)
    year_text = ''
    m = re.match(r'(\d{4})', str(year_quarter).strip())
    if m:
        year_text = f'{m.group(1)}年'

    warnings = []
    n_files = 0
    for po, ports in groups.items():
        first_recs = next(iter(ports.values()))
        style = str(first_recs[0].get('客款号') or '').strip()
        params = params_by_style.get(style)
        if not params or not params.get('profile'):
            warnings.append(f'{po}（款号 {style}）：未选包装方式/未填每箱上限，已跳过')
            continue
        wb = Workbook()
        wb.remove(wb.active)
        multi = len(ports) > 1
        for i, (port, recs) in enumerate(ports.items(), 1):
            ws = wb.create_sheet(f'{i}.0')
            po_disp = f'{po}-{i}' if multi else po
            _sd = (ship_dates or {}).get(f'{po}|{port}') or \
                  str(recs[0].get('实际出厂期') or '')
            write_packing_sheet(ws, po_disp, recs, params, year_text, maker, _sd,
                                params.get('tags'))
        if comp_pack:
            ws_meta = wb.create_sheet('_meta')
            ws_meta.sheet_state = 'hidden'
            ws_meta.append(['厂款号', '成分', '包装方式'])
            for s, cp in comp_pack.items():
                ws_meta.append([s, cp.get('comp', ''), cp.get('pack', '')])
        fac_disp = _fac_display(first_recs[0], params['profile'].pieces_per_pack)
        suffix = ''
        if not multi:
            port_cn = str(first_recs[0].get('港口') or '').strip()
            if port_cn:
                suffix = f'-{port_cn}'
        fname = sanitize_filename(f'{po} {style} {fac_disp} PL{suffix}.xlsx')
        try:
            wb.save(out / fname)
            n_files += 1
        except PermissionError:
            warnings.append(f'{fname}：文件被占用（可能正在 Excel 中打开），未能写入')
    return n_files, warnings


def scan_styles(po_folder, detail_folder):
    """Return OrderedDict 客款号 → {'fac': 厂款号 display, 'pos': [po...]}."""
    groups = build_po_groups(po_folder, detail_folder)
    styles = OrderedDict()
    for po, ports in groups.items():
        recs = next(iter(ports.values()))
        style = str(recs[0].get('客款号') or '').strip()
        ent = styles.setdefault(style, {'fac': _fac_display(recs[0]), 'pos': []})
        ent['pos'].append(po)
    return styles


# ══════════════════════════════════════════════════════════════
# 配置
# ══════════════════════════════════════════════════════════════

def load_cfg():
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text(encoding='utf-8'))
        except Exception:
            pass
    return {}


def save_cfg(**kw):
    cfg = load_cfg()
    cfg.update(kw)
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2),
                               encoding='utf-8')
    except Exception:
        pass


# ══════════════════════════════════════════════════════════════
# UI
# ══════════════════════════════════════════════════════════════

class InlineEditTreeview(ttk.Treeview):
    """Double-click editing; combo_cols={col_idx: values_provider} 用下拉框。"""

    def __init__(self, parent, editable_cols, combo_cols=None, **kw):
        super().__init__(parent, **kw)
        self.editable_cols = editable_cols
        self.combo_cols = combo_cols or {}
        self._widget = None
        self.bind('<Double-1>', self._start_edit)

    def _start_edit(self, event):
        if self.identify('region', event.x, event.y) != 'cell':
            return
        col_id = self.identify_column(event.x)
        row_id = self.identify_row(event.y)
        col_idx = int(col_id.replace('#', '')) - 1
        if col_idx not in self.editable_cols and col_idx not in self.combo_cols:
            return
        if self._widget:
            self._widget.destroy()
        bbox = self.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        values = self.item(row_id, 'values')

        def commit_value(new_val):
            vals = list(values)
            vals[col_idx] = new_val
            self.item(row_id, values=vals)
            if self._widget:
                self._widget.destroy()
                self._widget = None
            self.last_edit = (row_id, col_idx, new_val)
            self.event_generate('<<RowEdited>>')

        if col_idx in self.combo_cols:
            e = ttk.Combobox(self, font=('Arial', 9),
                             values=self.combo_cols[col_idx](), state='readonly')
            e.place(x=x, y=y, width=max(w, 160), height=h)
            e.set(values[col_idx])
            e.focus()
            e.bind('<<ComboboxSelected>>', lambda ev: commit_value(e.get()))
            e.bind('<Escape>',
                   lambda ev: (e.destroy(), setattr(self, '_widget', None)))
            self._widget = e
            return

        e = tk.Entry(self, font=('Arial', 9))
        e.place(x=x, y=y, width=w, height=h)
        e.insert(0, values[col_idx])
        e.select_range(0, 'end')
        e.focus()
        self._widget = e
        e.bind('<Return>', lambda ev: commit_value(e.get()))
        e.bind('<Tab>', lambda ev: commit_value(e.get()))
        e.bind('<FocusOut>', lambda ev: commit_value(e.get()))
        e.bind('<Escape>', lambda ev: (e.destroy(), setattr(self, '_widget', None)))


class ProfileManager(tk.Toplevel):
    """自定义包装方式管理（新建/修改/删除；内置配置只读，可另存为副本）。"""

    FIELDS = [
        ('name',      '名称'),
        ('tiers',     '大箱容量（件）'),
        ('cap_mid',   '中箱容量（件）'),
        ('cap_small', '小箱容量（件）'),
        ('allowance', '尾箱允许超出'),
        ('pieces',    '件套数(1=单件)'),
        ('net',       '单件净重kg'),
        ('w_big',     '大箱重kg'),
        ('w_mid',     '中箱重kg'),
        ('w_small',   '小箱重kg'),
    ]

    def __init__(self, master, cfg, on_change):
        super().__init__(master)
        self.title('包装方式配置')
        self.geometry('640x530')
        self.cfg = cfg
        self.on_change = on_change
        self.transient(master)

        left = tk.Frame(self); left.pack(side='left', fill='y', padx=8, pady=8)
        self.lb = tk.Listbox(left, width=30, height=18)
        self.lb.pack(fill='y', expand=True)
        self.lb.bind('<<ListboxSelect>>', self._on_select)

        right = tk.Frame(self); right.pack(side='left', fill='both',
                                           expand=True, padx=8, pady=8)
        self.vars = {}
        for r, (key, lbl) in enumerate(self.FIELDS):
            tk.Label(right, text=lbl + '：').grid(row=r, column=0, sticky='e', pady=3)
            v = tk.StringVar()
            self.vars[key] = v
            tk.Entry(right, textvariable=v, width=28).grid(row=r, column=1, sticky='w')
        tk.Label(right,
                 text='大箱容量填整数（如 90），或分级：2XS:90, L:80\n'
                      '选中内置配置后可修改值再点「保存」，名称相同则覆盖内置默认',
                 fg='gray', justify='left').grid(
            row=len(self.FIELDS), column=0, columnspan=2, sticky='w', pady=4)

        btns = tk.Frame(right)
        btns.grid(row=len(self.FIELDS) + 1, column=0, columnspan=2, pady=8)
        tk.Button(btns, text='保存', command=self._save).pack(side='left', padx=4)
        tk.Button(btns, text='删除 / 恢复默认', command=self._delete).pack(side='left', padx=4)
        tk.Button(btns, text='关闭', command=self.destroy).pack(side='left', padx=4)

        self._reload()

    def _profiles(self):
        return pp.all_profiles(self.cfg)

    def _reload(self):
        self.lb.delete(0, 'end')
        for p in self._profiles():
            tag = '' if p.id.startswith('custom:') else '【内置·只读】'
            self.lb.insert('end', f'{tag}{p.name_zh}')

    def _on_select(self, _evt=None):
        sel = self.lb.curselection()
        if not sel:
            return
        p = self._profiles()[sel[0]]
        self.vars['name'].set(p.name_zh)
        self.vars['tiers'].set(pp.tiers_to_text(p.tiers))
        self.vars['cap_mid'].set(str(p.cap_mid) if p.cap_mid else '')
        self.vars['cap_small'].set(str(p.cap_small) if p.cap_small else '')
        self.vars['allowance'].set(str(p.allowance))
        self.vars['pieces'].set(str(p.pieces_per_pack))
        self.vars['net'].set(str(p.net_k or ''))
        self.vars['w_big'].set(str(p.gross_add))
        self.vars['w_mid'].set(str(p.gross_add_mid))
        self.vars['w_small'].set(str(p.gross_add_small))

    def _save(self):
        name = self.vars['name'].get().strip()
        if not name:
            messagebox.showerror('错误', '请填写名称', parent=self); return
        try:
            tiers = pp.parse_tiers_text(self.vars['tiers'].get())
            if not tiers:
                raise ValueError('请填写每箱数量分级')

            def num(key, default):
                v = self.vars[key].get().strip()
                return float(v) if v else default
            prof = pp.PackProfile(
                id=f'custom:{name}',
                name_zh=name,
                tiers=tiers,
                cap_mid=int(num('cap_mid', 0)),
                cap_small=int(num('cap_small', 0)),
                net_k=num('net', 0.0),
                gross_add=num('w_big', 0.8),
                gross_add_mid=num('w_mid', 0.6),
                gross_add_small=num('w_small', 0.5),
                allowance=int(num('allowance', 3)),
                pieces_per_pack=int(num('pieces', 1)),
            )
        except Exception as ex:
            messagebox.showerror('错误', str(ex), parent=self); return
        customs = [p for p in pp.load_custom_profiles(self.cfg)
                   if p.id != prof.id]
        customs.append(prof)
        self.cfg['custom_pack_profiles'] = pp.dump_custom_profiles(customs)
        save_cfg(custom_pack_profiles=self.cfg['custom_pack_profiles'])
        self._reload()
        self.on_change()
        messagebox.showinfo('完成', f'已保存：{name}', parent=self)

    def _delete(self):
        sel = self.lb.curselection()
        if not sel:
            return
        p = self._profiles()[sel[0]]
        if not p.id.startswith('custom:'):
            messagebox.showinfo('提示', '这是内置默认配置，未被修改过，无需删除。\n'
                                '如需修改：先编辑字段，再点「保存」。', parent=self)
            return
        if not messagebox.askyesno('确认', f'删除「{p.name_zh}」？\n'
                                   '如果它覆盖了同名内置配置，删除后将恢复内置默认。',
                                   parent=self):
            return
        customs = [x for x in pp.load_custom_profiles(self.cfg) if x.id != p.id]
        self.cfg['custom_pack_profiles'] = pp.dump_custom_profiles(customs)
        save_cfg(custom_pack_profiles=self.cfg['custom_pack_profiles'])
        self._reload()
        self.on_change()


def build_into(parent):
    cfg = load_cfg()
    p = {'padx': 8, 'pady': 5}

    rows_cfg = [
        ('合同文件夹：',     'po_folder'),
        ('订单明细文件夹：', 'detail_folder'),
        ('输出文件夹：',     'out_dir'),
    ]
    parent.columnconfigure(1, weight=1)
    vars_ = {}
    for r, (lbl, key) in enumerate(rows_cfg):
        tk.Label(parent, text=lbl).grid(row=r, column=0, sticky='e', **p)
        v = tk.StringVar(value=cfg.get(key, ''))
        vars_[key] = v
        tk.Entry(parent, textvariable=v, width=20).grid(row=r, column=1, sticky='ew', **p)

        def mk_browse(var=v, k=key):
            def _b():
                path = filedialog.askdirectory()
                if path:
                    var.set(path)
                    save_cfg(**{k: path})
            return _b
        tk.Button(parent, text='文件夹…', command=mk_browse()).grid(row=r, column=2, **p)

    tk.Label(parent, text='年度/季度：').grid(row=3, column=0, sticky='e', **p)
    yq_var = tk.StringVar(value=cfg.get('year_quarter', ''))
    tk.Entry(parent, textvariable=yq_var, width=20).grid(row=3, column=1, sticky='w', **p)
    tk.Label(parent, text='制表人：').grid(row=4, column=0, sticky='e', **p)
    maker_var = tk.StringVar(value=cfg.get('maker', ''))
    tk.Entry(parent, textvariable=maker_var, width=20).grid(row=4, column=1, sticky='w', **p)

    btn_row = tk.Frame(parent)
    btn_row.grid(row=5, column=0, columnspan=3, sticky='w', **p)
    tk.Button(btn_row, text='① 扫描款号', width=14,
              command=lambda: do_scan()).pack(side='left')
    tk.Button(btn_row, text='包装方式配置…',
              command=lambda: open_manager()).pack(side='left', padx=(10, 0))

    def profile_names():
        return [p.name_zh for p in pp.all_profiles(cfg)]

    def open_manager():
        ProfileManager(parent.winfo_toplevel(), cfg, on_change=lambda: None)

    cols = ('style', 'fac', 'npo', 'profile', 'cap_big', 'cap_mid', 'cap_small', 'net',
            'w_big', 'w_mid', 'w_small',
            'tag_吊牌', 'tag_价钱牌', 'tag_衣架')

    # ── Notebook ─────────────────────────────────────────────────────
    nb_inner = ttk.Notebook(parent)
    nb_inner.grid(row=6, column=0, columnspan=3, sticky='nsew', padx=8, pady=4)
    parent.rowconfigure(6, weight=1)

    # ── Tab 1: 款号 & 装箱参数 ────────────────────────────────────────
    tab_pack = ttk.Frame(nb_inner)
    nb_inner.add(tab_pack, text='  款号 & 装箱参数  ')
    tab_pack.columnconfigure(0, weight=1)
    tab_pack.rowconfigure(0, weight=1)

    _tv_sb = ttk.Scrollbar(tab_pack, orient='vertical')
    tv = InlineEditTreeview(tab_pack, editable_cols=[4, 5, 6, 7, 8, 9, 10],
                            combo_cols={3: profile_names,
                                        11: lambda: ['√', ''],
                                        12: lambda: ['√', ''],
                                        13: lambda: ['√', '']},
                            columns=cols, show='headings',
                            yscrollcommand=_tv_sb.set)
    headers = ['客款号', '厂款号', 'PO数', '包装方式 ▾',
               '大箱qty ✏', '中箱qty ✏', '小箱qty ✏',
               '单件净重kg ✏', '大箱重kg ✏', '中箱重kg ✏', '小箱重kg ✏',
               '吊牌 ▾', '价钱牌 ▾', '衣架 ▾']
    widths = [85, 130, 45, 170, 75, 75, 75, 90, 80, 80, 80, 55, 60, 55]
    for c, h, w in zip(cols, headers, widths):
        tv.heading(c, text=h)
        tv.column(c, width=w, anchor='center')
    tv.grid(row=0, column=0, sticky='nsew')
    _tv_sb.config(command=tv.yview)
    _tv_sb.grid(row=0, column=1, sticky='ns')
    tk.Label(tab_pack,
             text='双击「包装方式」选配置；每箱上限可临时覆盖分级；留空用配置值',
             fg='gray').grid(row=1, column=0, columnspan=2, sticky='w', padx=4, pady=2)

    def fill_row_from_profile(row_id, prof):
        vals = list(tv.item(row_id, 'values'))
        while len(vals) < 14:
            vals.append('√')
        vals[4] = str(prof.tiers[0][1]) if prof.tiers else ''
        vals[5] = str(prof.cap_mid) if prof.cap_mid else ''
        vals[6] = str(prof.cap_small) if prof.cap_small else ''
        vals[7] = prof.net_k if prof.net_k else ''
        vals[8] = prof.gross_add
        vals[9] = prof.gross_add_mid
        vals[10] = prof.gross_add_small
        tv.item(row_id, values=vals)

    def on_row_edited(_evt=None):
        le = getattr(tv, 'last_edit', None)
        if not le:
            return
        row_id, col_idx, val = le
        if col_idx != 3:
            return
        prof = next((p for p in pp.all_profiles(cfg)
                     if p.name_zh == str(val).strip()), None)
        if prof:
            fill_row_from_profile(row_id, prof)

    tv.bind('<<RowEdited>>', on_row_edited)

    # ── Tab 2: 出货日期 ───────────────────────────────────────────────
    tab_date = ttk.Frame(nb_inner)
    nb_inner.add(tab_date, text='  出货日期  ')
    tab_date.columnconfigure(0, weight=1)
    tab_date.rowconfigure(1, weight=1)

    tk.Label(tab_date, text='出货日期（每合同/港口，默认取最终交货日，可覆盖）：',
             fg='#444').grid(row=0, column=0, columnspan=2, sticky='w', padx=6, pady=(6, 0))

    _date_inner = tk.Frame(tab_date)
    _date_inner.grid(row=1, column=0, columnspan=2, sticky='nsew', padx=4, pady=2)
    _date_inner.columnconfigure(0, weight=1)
    _date_inner.rowconfigure(1, weight=1)

    _COL_WIDTHS = [150, 110, 160]
    _ALL = '(全部)'
    _filter_frame = tk.Frame(_date_inner)
    _filter_frame.grid(row=0, column=0, sticky='ew')
    for ci, minsz in enumerate(_COL_WIDTHS):
        _filter_frame.columnconfigure(ci, minsize=minsz, weight=1 if ci == 2 else 0)
    _date_filter_vars = {}
    _date_filter_cbs  = {}
    for ci, col_key in enumerate(('po', 'port', 'date')):
        v = tk.StringVar(value=_ALL)
        _date_filter_vars[col_key] = v
        cb = ttk.Combobox(_filter_frame, textvariable=v, state='readonly', width=1)
        cb.grid(row=0, column=ci, sticky='ew', padx=(0, 1))
        _date_filter_cbs[col_key] = cb

    _date_sb = ttk.Scrollbar(_date_inner, orient='vertical')
    date_tv = InlineEditTreeview(_date_inner, editable_cols=[2],
                                 columns=('po', 'port', 'date'),
                                 show='headings',
                                 yscrollcommand=_date_sb.set)
    date_tv.heading('po',   text='合同号')
    date_tv.heading('port', text='港口')
    date_tv.heading('date', text='出货日期 ✏')
    date_tv.column('po',   width=150, anchor='center')
    date_tv.column('port', width=110, anchor='center')
    date_tv.column('date', width=160, anchor='center')
    date_tv.grid(row=1, column=0, sticky='nsew')
    _date_sb.config(command=date_tv.yview)
    _date_sb.grid(row=1, column=1, sticky='ns')

    _date_all_rows = []   # master list: [(po, port, date), ...] — includes filtered-out rows

    def _sync_date_edits():
        visible = {(v[0], v[1]): v[2]
                   for item in date_tv.get_children()
                   for v in [date_tv.item(item, 'values')]}
        for i, (po, port, _) in enumerate(_date_all_rows):
            if (po, port) in visible:
                _date_all_rows[i] = (po, port, visible[(po, port)])

    def _refresh_filter_options():
        for ci, col_key in enumerate(('po', 'port', 'date')):
            unique = sorted(set(row[ci] for row in _date_all_rows if row[ci]))
            _date_filter_cbs[col_key]['values'] = [_ALL] + unique

    def _apply_date_filter(*_):
        _sync_date_edits()
        active = {k: v.get() for k, v in _date_filter_vars.items()}
        for item in date_tv.get_children():
            date_tv.delete(item)
        for po, port, date in _date_all_rows:
            row_map = {'po': po, 'port': port, 'date': date}
            if all(sel == _ALL or sel == row_map[k] for k, sel in active.items()):
                date_tv.insert('', 'end', values=(po, port, date))

    for v in _date_filter_vars.values():
        v.trace_add('write', _apply_date_filter)

    # ── Tab 3: 成分 & 包装方式 ────────────────────────────────────────
    tab_cp = ttk.Frame(nb_inner)
    nb_inner.add(tab_cp, text='  成分 & 包装方式  ')
    tab_cp.columnconfigure(0, weight=1)
    tab_cp.rowconfigure(0, weight=1)

    cp_cols = ('style', 'comp', 'pack')
    cp_tv = InlineEditTreeview(tab_cp, editable_cols=[1, 2],
                               columns=cp_cols, show='headings')
    cp_tv.heading('style', text='厂款号')
    cp_tv.heading('comp',  text='成分 ✏')
    cp_tv.heading('pack',  text='包装方式 ✏')
    cp_tv.column('style', width=110, anchor='center')
    cp_tv.column('comp',  width=400, anchor='w')
    cp_tv.column('pack',  width=200, anchor='w')
    cp_sb = ttk.Scrollbar(tab_cp, orient='vertical', command=cp_tv.yview)
    cp_tv.configure(yscrollcommand=cp_sb.set)
    cp_tv.grid(row=0, column=0, sticky='nsew')
    cp_sb.grid(row=0, column=1, sticky='ns')
    tk.Label(tab_cp, text='  双击单元格直接编辑', fg='gray').grid(
        row=1, column=0, sticky='w', padx=4, pady=2)

    comp_pack_data: dict[str, dict] = {}
    _cp_saved: dict = dict(cfg.get('comp_pack_by_style', {}))

    def _refresh_cp_tv():
        cp_tv.delete(*cp_tv.get_children())
        for style in sorted(comp_pack_data.keys()):
            d = comp_pack_data[style]
            cp_tv.insert('', 'end', iid=style,
                         values=(style, d.get('comp', ''), d.get('pack', '')))

    def _on_cp_edited(_evt=None):
        le = getattr(cp_tv, 'last_edit', None)
        if not le:
            return
        row_id, col_idx, val = le
        if col_idx == 1:
            comp_pack_data.setdefault(row_id, {})['comp'] = val.strip()
        elif col_idx == 2:
            comp_pack_data.setdefault(row_id, {})['pack'] = val.strip()
        save_cfg(comp_pack_by_style=dict(comp_pack_data))

    cp_tv.bind('<<RowEdited>>', _on_cp_edited)

    status = tk.StringVar(value='')
    tk.Label(parent, textvariable=status, fg='#005599').grid(
        row=7, column=0, columnspan=3, sticky='w', **p)

    def saved_style_params():
        return cfg.get('style_params', {})

    def do_scan():
        po = vars_['po_folder'].get().strip()
        det = vars_['detail_folder'].get().strip()
        if not po:
            messagebox.showerror('错误', '请先选择合同文件夹'); return
        status.set('扫描中…')
        parent.update_idletasks()
        try:
            groups = build_po_groups(po, det or None)
            styles = scan_styles(po, det or None)
        except Exception as ex:
            status.set('')
            messagebox.showerror('扫描失败', str(ex)); return
        for item in tv.get_children():
            tv.delete(item)
        saved_dates = cfg.get('po_dates', {})
        _date_all_rows.clear()
        for po_key, ports in groups.items():
            for port_key, recs in ports.items():
                default_date = _fmt_date(recs[0].get('实际出厂期') or '')
                saved = saved_dates.get(f'{po_key}|{port_key}', '')
                _date_all_rows.append((po_key, port_key, saved or default_date))
        _refresh_filter_options()
        _apply_date_filter()
        remembered = saved_style_params()
        profile_by_id = {p.id: p for p in pp.all_profiles(cfg)}
        for style, info in styles.items():
            sp = remembered.get(style, {})
            prof = profile_by_id.get(sp.get('profile_id', ''))
            tv.insert('', 'end', values=(
                style, info['fac'], len(info['pos']),
                prof.name_zh if prof else '',
                sp.get('cap_big', ''), sp.get('cap_mid', ''), sp.get('cap_small', ''),
                sp.get('net', ''),
                sp.get('w_big', ''), sp.get('w_mid', ''), sp.get('w_small', ''),
                sp.get('tag_吊牌', '√'), sp.get('tag_价钱牌', '√'), sp.get('tag_衣架', '√')))
        for style, info in styles.items():
            fac = info['fac']
            if fac not in comp_pack_data:
                saved = _cp_saved.get(fac, {})
                comp_pack_data[fac] = {'comp': saved.get('comp', ''),
                                       'pack': saved.get('pack', '')}
        _refresh_cp_tv()
        status.set(f'扫描到 {len(styles)} 个款号；请选择包装方式并填写大/中/小箱容量')

    def collect_params():
        """Returns ({style: params}, missing, {style: saved_dict})."""
        out, missing, save_map = {}, [], {}
        profile_by_name = {p.name_zh: p for p in pp.all_profiles(cfg)}

        def num(v):
            try:
                return float(v)
            except (TypeError, ValueError):
                return None

        def _safe_int(v):
            try:
                return int(float(v)) if str(v).strip() else None
            except (TypeError, ValueError):
                return None

        for item in tv.get_children():
            style, fac, npo, prof_name, cap_big, cap_mid, cap_small, net, wb_, wm, ws_, \
                tag_吊牌, tag_价钱牌, tag_衣架 = tv.item(item, 'values')
            prof = profile_by_name.get(str(prof_name).strip())
            cap_big_int = _safe_int(cap_big)
            if prof is None and cap_big_int:
                prof = pp.PackProfile(id='custom:临时', name_zh='临时',
                                      tiers=(('*', cap_big_int),))
            if prof is None:
                missing.append(style)
                continue
            out[style] = {
                'profile': prof,
                'cap_big': cap_big_int,
                'cap_mid': _safe_int(cap_mid),
                'cap_small': _safe_int(cap_small),
                'net': num(net),
                'w_big': num(wb_),
                'w_mid': num(wm),
                'w_small': num(ws_),
                'tags': {'吊牌': tag_吊牌 == '√',
                         '价钱牌': tag_价钱牌 == '√',
                         '衣架': tag_衣架 == '√'},
            }
            save_map[style] = {
                'profile_id': prof.id if not prof.id.startswith('custom:临时') else '',
                'cap_big': cap_big, 'cap_mid': cap_mid, 'cap_small': cap_small,
                'net': net,
                'w_big': wb_, 'w_mid': wm, 'w_small': ws_,
                'tag_吊牌': tag_吊牌, 'tag_价钱牌': tag_价钱牌, 'tag_衣架': tag_衣架,
            }
        return out, missing, save_map

    def do_generate():
        po        = vars_['po_folder'].get().strip()
        det       = vars_['detail_folder'].get().strip()
        out       = vars_['out_dir'].get().strip()
        yq        = yq_var.get().strip()
        _sync_date_edits()
        ship_dates = {f'{po}|{port}': date.strip()
                      for po, port, date in _date_all_rows}
        if not po:
            messagebox.showerror('错误', '请先选择合同文件夹'); return
        if not out:
            messagebox.showerror('错误', '请指定输出文件夹'); return
        if not tv.get_children():
            messagebox.showerror('错误', '请先点击「扫描款号」并填写装箱参数'); return
        params, missing, save_map = collect_params()
        if missing:
            if not messagebox.askyesno(
                    '缺少参数',
                    '以下款号未选包装方式（也未填每箱上限），对应 PO 将跳过：\n\n'
                    + '\n'.join(missing) + '\n\n是否继续？'):
                return
        # remember
        remembered = saved_style_params()
        remembered.update(save_map)
        _cp_saved.update(comp_pack_data)
        save_cfg(po_folder=po, detail_folder=det, out_dir=out,
                 year_quarter=yq, maker=maker_var.get().strip(),
                 po_dates=ship_dates, style_params=remembered,
                 comp_pack_by_style=dict(comp_pack_data))
        cfg['style_params'] = remembered

        gen_btn.config(state='disabled')
        status.set('生成中…')
        result = {}

        def worker():
            try:
                result['n'], result['warn'] = generate_packing_lists(
                    po, det or None, out, yq, maker_var.get().strip(), params, ship_dates,
                    comp_pack=dict(comp_pack_data))
            except Exception:
                result['err'] = traceback.format_exc()

        t = threading.Thread(target=worker, daemon=True)
        t.start()

        def poll():
            if t.is_alive():
                parent.after(150, poll)
                return
            gen_btn.config(state='normal')
            if 'err' in result:
                status.set('生成失败')
                messagebox.showerror('生成失败', result['err'])
                return
            status.set(f'完成：{result["n"]} 个装箱单')
            msg = f'已生成 {result["n"]} 个装箱单：\n{out}'
            if result['warn']:
                msg += '\n\n警告：\n' + '\n'.join(result['warn'])
            messagebox.showinfo('完成', msg)

        parent.after(150, poll)

    gen_btn = tk.Button(parent, text='  ② 生成装箱单  ', command=do_generate,
                        bg='#1565C0', fg='white', font=('Arial', 11, 'bold'))
    gen_btn.grid(row=8, column=0, columnspan=3, pady=10)


def main():
    root = tk.Tk()
    root.title('装箱单生成工具（测试版）')
    root.geometry('980x820')
    build_into(root)
    root.mainloop()


if __name__ == '__main__':
    main()
