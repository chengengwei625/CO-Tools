"""ASOS PO + existing factory packing sheet -> checked carton labels and ASN packing list."""
import json
import math
import re
import sys
import tempfile
from collections import Counter
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import fitz
import openpyxl
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

# Same font and size as the ASOS Carton Label Template barcode cells (Free 3 of 9 Extended, 36pt).
BARCODE_FONT, BARCODE_SIZE = 'Free3of9Extended', 36
BARCODE_FILE = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent)) / 'fre3of9x.ttf'
# Saved settings live next to the exe, like the other tabs' *_config.json.
CONFIG_FILE = (Path(sys.executable).parent if getattr(sys, 'frozen', False)
               else Path(__file__).parent) / 'asos_config.json'


def load_cfg():
    try:
        return json.loads(CONFIG_FILE.read_text(encoding='utf-8'))
    except Exception:  # missing or unreadable config: start empty
        return {}


def save_cfg(**kw):
    cfg = load_cfg()
    cfg.update(kw)
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding='utf-8')
    except Exception:  # read-only folder: settings just aren't remembered
        pass
# Main template has SKU A–C; the "Extra SKUs" template has SKU A–D.
# ponytail: 5–8 SKU cartons blocked until ASOS's 8-SKU template is supplied.
MAX_SKUS = 4


def positive_int(value, field):
    if isinstance(value, bool) or not re.fullmatch(r"[0-9]+(?:\.0+)?", str(value).strip()):
        raise ValueError(f"{field} 必须是正整数：{value}")
    result = int(float(value))
    if result <= 0:
        raise ValueError(f"{field} 必须大于 0")
    return result


def normalize_asn(value):
    value = value.strip()
    if not re.fullmatch(r"[0-9]{12,14}", value):
        raise ValueError("ASN 必须为 12–14 位数字；不足 14 位自动在前面补零")
    return value.zfill(14)


class NotAsosPo(ValueError):
    """PDF is not an ASOS Purchase Order (batch folders may hold other PDFs)."""


PO_COLUMNS = {'SKU ID': 'sku', 'Brand': 'brand', 'Brand Size': 'size', 'First Dest': 'first_destination',
              'Final Dest': 'destination', 'Units': 'qty', 'Packing Method': 'packing',
              'Website Color': 'colour', 'Option ID': 'option'}


def po_table_rows(page):
    """SKU rows of one PO page, read by position under the bordered header cells.
    Data rows have no borders, and Barcode may be blank or wrap to two lines, so rows are
    anchored on the SKU ID cell. The PDF page is stored rotated; words are mapped to page view."""
    for table in page.find_tables().tables:
        found = [(names, row) for names, row in zip(table.extract(), table.rows)
                 if 'SKU ID' in [' '.join((n or '').split()) for n in names]]
        if found:
            names, header = found[0]
            break
    else:
        return []
    columns = [(PO_COLUMNS[' '.join((n or '').split())], cell[0], cell[2])
               for n, cell in zip(names, header.cells) if cell and ' '.join((n or '').split()) in PO_COLUMNS]
    top = header.bbox[3]
    words = []
    for w in page.get_text('words'):
        r = fitz.Rect(w[:4]) * page.rotation_matrix
        words.append(((r.x0 + r.x1) / 2, r.y0, w[4]))
    # Footer ("Print Date: …", "Page 2 of 2") sits under the table inside the column ranges.
    footer = min([y for _, y, t in words if t in ('Print', 'Page') and y > top], default=float('inf'))
    cells = {}
    for x, y, text in words:
        if top < y < footer:
            for key, x0, x1 in columns:
                if x0 <= x < x1:
                    cells.setdefault(key, []).append((y, text))
    anchors = sorted(y for y, _ in cells.get('sku', []))
    rows = [{} for _ in anchors]
    for key, items in cells.items():
        for y, text in sorted(items):
            index = max((i for i, a in enumerate(anchors) if a <= y + 3), default=None)
            if index is None:
                raise ValueError(f"PO 表格中“{text}”不在任何 SKU 行内，无法解析")
            rows[index][key] = (rows[index].get(key, '') + ' ' + text).strip()
    return rows


def read_asos_po(path):
    with fitz.open(str(path)) as doc:
        text = "\n".join(page.get_text() for page in doc)
        po_ids = set(re.findall(r"PO ID:\s*([0-9]+)", text))
        if len(po_ids) != 1 or "ASOS Purchase Order" not in text:
            raise NotAsosPo("请选择一份文字型 ASOS Purchase Order PDF（只允许一个 PO ID）")
        po_id = po_ids.pop()
        try:
            table = [row for page in doc for row in po_table_rows(page)]
        except ValueError as exc:
            exc.po_id = po_id
            raise
    try:
        rows = []
        for number, r in enumerate(table, 1):
            missing = [name for name, key in PO_COLUMNS.items() if not r.get(key)]
            if missing or not r['sku'].isdigit():
                raise ValueError(f"PO 表格第 {number} 行无法读取（缺少 {', '.join(missing) or 'SKU ID'}）")
            rows.append(dict(size=r['size'].upper(), sku=r['sku'], qty=positive_int(r['qty'], "PO 数量"),
                             colour=r['colour'].upper(), destination=r['destination'],
                             first_destination=r['first_destination'], option=r['option'],
                             brand=r['brand'], packing=r['packing']))
        if not rows:
            raise ValueError("未能识别 ASOS SKU 表格")
        if len({(r['colour'], r['option'], r['destination'], r['first_destination']) for r in rows}) != 1:
            raise ValueError("一份任务只支持一个 ASOS 款色及目的仓；请先拆分 PO")
        if len({r['size'] for r in rows}) != len(rows) or len({r['sku'] for r in rows}) != len(rows):
            raise ValueError("PO 中存在重复尺码或 SKU，不能自动匹配")
        if any(r['packing'].lower() != 'flat' for r in rows):
            raise ValueError("当前仅支持 Flat 纸箱包装，不能生成挂装标签")
        total = re.search(r"Total Units:\s*([0-9]+)|^([0-9]+)\s*\nTotal Units:", text, re.M)
        if total is None or sum(r['qty'] for r in rows) != int(total.group(1) or total.group(2)):
            raise ValueError("PO SKU 明细合计与 Total Units 不符，或无法读取总数")
    except ValueError as exc:
        exc.po_id = po_id
        raise
    return dict(po=po_id, rows=rows, colour=rows[0]['colour'], destination=rows[0]['destination'])


def read_packing(path):
    """Read the existing A..AI factory sheet without relying on formula caches."""
    workbook = openpyxl.load_workbook(path, data_only=False)
    try:
        sheets = [s for s in workbook if str(s['A1'].value or '').strip() == '装箱资料表']
        if len(sheets) != 1:
            raise ValueError("装箱文件必须含一张“装箱资料表”；请按订单分别导出")
        sheet = sheets[0]
        if str(sheet['B12'].value or '').strip() != '箱号':
            raise ValueError("无法识别装箱单表头（需要本程序的装箱资料表格式）")
        sizes = {col: str(sheet.cell(12, col).value).strip().upper()
                 for col in range(6, 25) if sheet.cell(12, col).value is not None}
        if len(set(sizes.values())) != len(sizes):
            raise ValueError("装箱单尺码表头重复")
        groups, group, used = [], None, set()
        for row in range(14, sheet.max_row + 1):
            if str(sheet.cell(row, 4).value or '').strip() == '共':
                break
            start, end = sheet.cell(row, 2).value, sheet.cell(row, 4).value
            colour = str(sheet.cell(row, 1).value or '').strip()
            if start is None and end is None:
                if colour and group is not None:
                    group['names'].append(colour)
                else:
                    group = None
                continue
            start, end = positive_int(start, f"第 {row} 行起始箱号"), positive_int(end, f"第 {row} 行结束箱号")
            if end < start or end - start > 9999:
                raise ValueError(f"第 {row} 行箱号范围不正确")
            if group is None:
                if not colour:
                    raise ValueError(f"第 {row} 行缺少颜色")
                group = dict(names=[], cartons=[])
                groups.append(group)
            if colour:
                group['names'].append(colour)
            quantities = {}
            for col, size in sizes.items():
                value = sheet.cell(row, col).value
                if value is not None and value != 0:
                    quantities[size] = positive_int(value, f"第 {row} 行 {size} 数量")
            if not quantities:
                raise ValueError(f"第 {row} 行没有有效数量")
            if any(sheet.cell(row, col).value not in (None, 0) for col in range(6, 25) if col not in sizes):
                raise ValueError(f"第 {row} 行存在没有尺码表头的数量")
            for number in range(start, end + 1):
                if number in used:
                    raise ValueError(f"装箱单箱号 {number} 重复")
                used.add(number)
                group['cartons'].append(dict(source=number, sizes=dict(quantities)))
        if not groups:
            raise ValueError("装箱单没有可读取的箱明细")
        destination = str(sheet['G2'].value or '') + ' ' + str(sheet['G3'].value or '')
        if 'ASOS' not in destination.upper():
            raise ValueError("所选装箱单不是 ASOS 订单")
        return dict(internal_po=str(sheet['D3'].value or ''), destination=destination,
                    style=str(sheet['D5'].value or ''), groups=groups)
    finally:
        workbook.close()


def destination_matches(packing, po):
    """Packing-list market (UK/EU/USA in 目的地) must match the PO's ASOS warehouse."""
    return not any(re.search(r'\b' + market + r'\b', packing['destination'].upper()) and po['destination'] != warehouse
                   for market, warehouse in [('UK', 'FC01'), ('EU', 'FC04'), ('USA', 'FC03')])


def parse_asn_mapping(text):
    """Typed lines 'ASOS PO ID  ASN  COG PO' (COG PO anywhere on the line, e.g. PO159521)
    -> {po_id: {'asn': 14-digit ASN, 'cog': 'PO159521'}}."""
    mapping = {}
    for number, line in enumerate(text.splitlines(), 1):
        tokens = line.replace(',', ' ').upper().split()
        if not tokens:
            continue
        cogs = [t for t in tokens if re.fullmatch(r'PO[0-9]+', t)]
        digits = [t for t in tokens if t.isdigit()]
        if len(cogs) != 1 or len(digits) != 2 or len(tokens) != 3:
            raise ValueError(f"ASN 对照第 {number} 行应为：ASOS PO ID  ASN  合同号（如 PO159521）")
        po_id, asn = digits
        if po_id in mapping:
            raise ValueError(f"ASN 对照中 ASOS PO {po_id} 重复")
        try:
            mapping[po_id] = dict(asn=normalize_asn(asn), cog=cogs[0])
        except ValueError as exc:
            raise ValueError(f"ASN 对照第 {number} 行：{exc}") from None
    if not mapping:
        raise ValueError("请填写 ASN 对照（每行：ASOS PO ID  ASN  合同号）")
    if len({m['asn'] for m in mapping.values()}) != len(mapping):
        raise ValueError("ASN 对照中同一 ASN 对应了多个 ASOS PO")
    return mapping


def load_packings(paths, cogs):
    """Open only packing lists whose file name contains a needed 合同号; D3 must confirm it.
    Returns ([(file, packing)], [skip notes], ignored count)."""
    packings, skipped, ignored = [], [], 0
    for path in sorted(set(paths)):
        name = Path(path).name
        if name.startswith('~$') or not any(re.search(c + r'(?![0-9])', name.upper()) for c in cogs):
            ignored += 1  # unrelated packing list (or Excel lock file)
            continue
        try:
            packing = read_packing(path)
        except Exception as exc:
            skipped.append(f"{name}：{exc}")
            continue
        if packing['internal_po'].strip().upper() not in cogs:
            skipped.append(f"{name}：合同号 {packing['internal_po']} 与文件名不符")
            continue
        packings.append((name, packing))
    return packings, skipped, ignored


def plan_batch(po_paths, packing_paths, mapping_text, supplier, description, ncs=False):
    """Match each PO PDF to its ASN and packing colour group; failures are reported per PO."""
    if not po_paths or not packing_paths:
        raise ValueError("请选择 ASOS PO PDF 及装箱资料表")
    mapping = parse_asn_mapping(mapping_text)
    packings, skipped, ignored = load_packings(packing_paths, {m['cog'] for m in mapping.values()})
    jobs, errors, used, seen, other_pdfs = [], [], {}, set(), 0
    for path in sorted(set(po_paths)):
        try:
            po = read_asos_po(path)
            if po['po'] in seen:
                raise ValueError(f"ASOS PO {po['po']} 的 PDF 重复")
            seen.add(po['po'])
            if po['po'] not in mapping:
                raise ValueError(f"ASN 对照中没有 ASOS PO {po['po']}")
            cog = mapping[po['po']]['cog']
            files = [(file, packing) for file, packing in packings if packing['internal_po'].strip().upper() == cog]
            if not files:
                raise ValueError(f"所选装箱单中找不到合同号 {cog}（文件名需包含合同号）")
            matches = [(file, packing, index) for file, packing in files
                       for index, group in enumerate(packing['groups'])
                       if po['colour'] in {n.upper() for n in group['names']} and destination_matches(packing, po)]
            only = [(file, packing, 0) for file, packing in files
                    if len(packing['groups']) == 1 and destination_matches(packing, po)]
            note = ''
            if not matches and len(only) == 1:
                # e.g. PO "MULTI" vs packing "RED DITSY/HEART": the 合同号 already pins the order,
                # and its only colour group still has to pass the quantity check.
                matches = only
                note = (f"颜色名称不同：PO {po['colour']} / 装箱单 {' / '.join(only[0][1]['groups'][0]['names'])}"
                        f"（{cog} 只有这一个颜色组，请确认）")
            if not matches:
                raise ValueError(f"{cog} 装箱单中找不到 {po['colour']} / {po['destination']} 的颜色组")
            if len(matches) > 1:
                raise ValueError(f"{cog} 有多个 {po['colour']} / {po['destination']} 颜色组，请只保留正确版本："
                                 + '、'.join(sorted({file for file, _, _ in matches})))
            file, packing, index = matches[0]
            if (file, index) in used:
                raise ValueError(f"与 ASOS PO {used[file, index]} 匹配到同一颜色组")
            job = make_job(po, packing, index, mapping[po['po']]['asn'], supplier, description, bool(note))
            used[file, index] = po['po']
            jobs.append(dict(job, packing_file=file, note=note))
        except NotAsosPo:
            other_pdfs += 1
        except Exception as exc:
            seen.add(getattr(exc, 'po_id', None))  # a selected-but-unreadable PO is not "unused"
            errors.append(f"{Path(path).name}：{exc}")
    assign_marks(jobs, ncs)
    return dict(jobs=jobs, errors=errors, unused=sorted(set(mapping) - seen),
                skipped=skipped, ignored=ignored, other_pdfs=other_pdfs)


def make_job(po, packing, group_index, asn, supplier, description, any_colour=False):
    asn = normalize_asn(asn)
    supplier, description = supplier.strip(), description.strip()
    for field, value, limit in [('供应商', supplier, 65), ('货物描述', description, 80)]:
        if not value or len(value) > limit or not value.isascii() or any(ord(c) < 32 for c in value):
            raise ValueError(f"{field}请填写英文（1–{limit} 个字符）")
    group = packing['groups'][group_index]
    if not destination_matches(packing, po):
        raise ValueError(f"装箱目的地 {packing['destination'].strip()} 与 ASOS 目的仓 {po['destination']} 不匹配")
    if not any_colour and po['colour'] not in {n.upper() for n in group['names']}:
        raise ValueError(f"所选装箱颜色与 ASOS PO 的 {po['colour']} 不匹配；请核对装箱单英文颜色")
    totals = Counter()
    for carton in group['cartons']:
        totals.update(carton['sizes'])
    expected = {r['size']: r['qty'] for r in po['rows']}
    if dict(totals) != expected:
        details = ', '.join(f"{s}: 装箱 {totals.get(s, 0)} / PO {expected.get(s, 0)}"
                            for s in sorted(set(totals) | set(expected)) if totals.get(s, 0) != expected.get(s, 0))
        raise ValueError("尺码数量不符，已阻止生成：" + details)
    by_size = {r['size']: r for r in po['rows']}
    mixed_counts = Counter()
    cartons = []
    for number, carton in enumerate(group['cartons'], 1):
        items = [dict(sku=by_size[s]['sku'], size=s, qty=q) for s, q in carton['sizes'].items()]
        if len(items) > MAX_SKUS:
            raise ValueError(f"原箱 {carton['source']} 有 {len(items)} 个 SKU；ASOS 标签模板最多 {MAX_SKUS} 个，请调整装箱")
        if len(items) > 1:
            mixed_counts.update(item['sku'] for item in items)
        cartons.append(dict(number=number, source=carton['source'], items=items))
    if any(count > 2 for count in mixed_counts.values()):
        raise ValueError("同一 SKU 出现在超过两只混码箱中，请先调整装箱")
    return dict(po=po['po'], destination=po['destination'], colour=po['colour'],
                asn=asn, supplier=supplier, description=description, cartons=cartons,
                internal_po=packing['internal_po'], style=packing['style'])


def draw_text(pdf, text, x, y, width, size=13):
    text = str(text)
    pdf.setFont('Helvetica', size)
    if pdf.stringWidth(text, 'Helvetica', size) > width:
        size *= width / pdf.stringWidth(text, 'Helvetica', size)
        pdf.setFont('Helvetica', size)
    pdf.drawString(x, y, text)


BAR_HEIGHT = 27  # bars of Free 3 of 9 Extended at 36pt (font ascent 0.752 em)


def draw_barcode(pdf, value, x0, x1, bottom, height):
    """Draw '*value*' in the template's barcode font, centred in the cell."""
    text = f"*{value}*"
    text_width = pdf.stringWidth(text, BARCODE_FONT, BARCODE_SIZE)
    # Never shrink a barcode to fit: oversized data needs a different approved layout.
    if text_width > x1 - x0 - 6:
        raise ValueError("条码超出标签宽度，不能缩小打印；请核对 ASN / SKU / 数量")
    pdf.setFont(BARCODE_FONT, BARCODE_SIZE)
    pdf.drawString((x0 + x1 - text_width) / 2, bottom + (height - BAR_HEIGHT) / 2, text)


def cell_text(pdf, text, x0, x1, bottom, height, size):
    """Centre text in a label cell like the template; shrink only if it would overflow the cell."""
    text = str(text)
    size = min(size, size * (x1 - x0 - 12) / max(pdf.stringWidth(text, 'Helvetica', size), 1))
    pdf.setFont('Helvetica', size)
    pdf.drawCentredString((x0 + x1) / 2, bottom + height / 2 - size * 0.35, text)


# Row heights (pt) and font sizes copied from ASOS_Carton Label Template.xlsx (SKU A–C) and
# Shipipng Label Extra SKUs.xlsx (SKU A–D). Rows shrink evenly if a layout is taller than the
# page; fonts and barcodes never do.
LABEL_LAYOUTS = {
    3: (dict(ac=35.25, from_=39.75, po=39.75, asn_bar=38.5, asn_text=35, head=18, sku_bar=45, sku_text=45,
             combined=38.5, ctn=35, skus_bar=40.5, skus_text=40.5),
        dict(ac=16, title=24, value=16, asn=24, head=16, sku=24, qty=36, ctn_title=11, ctn=24,
             skus_title=11, skus=26)),
    4: (dict(ac=21, from_=30, po=30, asn_bar=44.25, asn_text=37.5, head=19.5, sku_bar=44.25, sku_text=37.5,
             combined=44.25, ctn=30, skus_bar=44.25, skus_text=30),
        dict(ac=16, title=24, value=16, asn=24, head=16, sku=24, qty=22, ctn_title=16, ctn=24,
             skus_title=16, skus=24)),
}


COLOURS = {'Red': colors.red, 'Blue': colors.blue, 'Green': colors.green,
           'Yellow': colors.yellow, 'Orange': colors.orange, 'Purple': colors.purple,
           'NCS - no colour': None}
MARK_SHAPES = {'block': '方块', 'circle': '圆形', 'triangle': '三角形', 'star': '星形'}


def assign_marks(jobs, ncs=False):
    """Every ASN in the batch gets its own colour/shape mark (manual p.12: each ASN different;
    use symbols once colours run out): 6 colour blocks, then circles, triangles, stars."""
    palette = [name for name, value in COLOURS.items() if value is not None]
    marks = [(colour, shape) for shape in MARK_SHAPES for colour in palette]
    if not ncs and len(jobs) > len(marks):
        raise ValueError(f"本批 {len(jobs)} 个 ASN 超过 {len(marks)} 种颜色/形状标记，请分批生成或使用 NCS")
    for index, job in enumerate(jobs):
        job['mark'], job['shape'] = ('NCS - no colour', None) if ncs else marks[index]


def mark_name(job):
    return 'NCS' if job['shape'] is None else f"{job['mark']} {MARK_SHAPES[job['shape']]}"


def draw_mark(pdf, colour, shape, cx, cy, size):
    """ASN colour code in the A/C cell: the template's coloured block, or a coloured symbol."""
    pdf.setFillColor(COLOURS[colour])
    r = size / 2
    if shape == 'circle':
        pdf.circle(cx, cy, r, stroke=0, fill=1)
    elif shape in ('triangle', 'star'):
        points = 3 if shape == 'triangle' else 10  # star: alternating outer/inner points
        path = pdf.beginPath()
        for i in range(points):
            angle = math.pi / 2 + 2 * math.pi * i / points
            radius = r if shape == 'triangle' or i % 2 == 0 else r * 0.45
            (path.moveTo if i == 0 else path.lineTo)(cx + radius * math.cos(angle), cy + radius * math.sin(angle))
        path.close()
        pdf.drawPath(path, stroke=0, fill=1)
    else:
        pdf.rect(cx - 25, cy - size / 2, 50, size, stroke=0, fill=1)
    pdf.setFillColor(colors.black)


def write_labels(job, path, colour, shape='block'):
    if not BARCODE_FILE.is_file():
        raise ValueError(f"缺少条码字体文件：{BARCODE_FILE}")
    pdfmetrics.registerFont(TTFont(BARCODE_FONT, str(BARCODE_FILE)))
    pdf = canvas.Canvas(str(path), pagesize=A4)
    x, right, top, bottom = 25, A4[0] - 25, A4[1] - 20, 20
    middle = x + (right - x) * 222 / 466.5  # template column widths C : D
    left_cell, right_cell, full = (x, middle), (middle, right), (x, right)
    for carton in job['cartons']:
        items = carton['items']
        if len(items) > MAX_SKUS:  # make_job blocks this; never drop SKUs silently
            raise ValueError(f"ASOS 箱 {carton['number']} 超过 {MAX_SKUS} 个 SKU")
        slots = max(3, len(items))  # 4 SKUs -> Extra SKUs template layout
        rows, font = LABEL_LAYOUTS[slots]
        total = (sum(rows[k] for k in ('ac', 'from_', 'po', 'asn_bar', 'asn_text', 'ctn', 'skus_bar', 'skus_text'))
                 + slots * sum(rows[k] for k in ('head', 'sku_bar', 'sku_text', 'combined')))
        h = {k: v * min(1, (top - bottom) / total) for k, v in rows.items()}
        y = top

        def rule(row_bottom, split=True):
            pdf.line(x, row_bottom, right, row_bottom)
            if split:
                pdf.line(middle, y, middle, row_bottom)

        def shade(x0, x1, row_bottom, height):  # grey cells as in the template
            pdf.setFillColor(colors.HexColor('#d9d9d9'))
            pdf.rect(x0, row_bottom, x1 - x0, height, stroke=0, fill=1)
            pdf.setFillColor(colors.black)

        b = y - h['ac']
        shade(*full, b, h['ac'])
        rule(b)
        cell_text(pdf, 'A/C', *left_cell, b, h['ac'], font['ac'])
        if COLOURS[colour] is not None:
            draw_mark(pdf, colour, shape, (middle + right) / 2, b + h['ac'] / 2, min(h['ac'] - 4, 22))
        y = b
        for title, value, key in [('FROM', job['supplier'], 'from_'), ('PO NUMBER', job['po'], 'po')]:
            b = y - h[key]
            rule(b)
            cell_text(pdf, title, *left_cell, b, h[key], font['title'])
            cell_text(pdf, value, *right_cell, b, h[key], font['value'])
            y = b
        b = y - h['asn_bar'] - h['asn_text']  # ASN title spans the barcode and number rows
        rule(b)
        pdf.line(middle, b + h['asn_text'], right, b + h['asn_text'])
        cell_text(pdf, 'ASN', *left_cell, b, y - b, font['title'])
        draw_barcode(pdf, job['asn'], *right_cell, b + h['asn_text'], h['asn_bar'])
        cell_text(pdf, job['asn'], *right_cell, b, h['asn_text'], font['asn'])
        y = b
        for slot in range(slots):
            b = y - h['head']
            shade(*full, b, h['head'])
            rule(b)
            cell_text(pdf, 'SKU ' + chr(65 + slot), *left_cell, b, h['head'], font['head'])
            cell_text(pdf, 'QTY ' + chr(65 + slot), *right_cell, b, h['head'], font['head'])
            y = b
            b = y - h['sku_bar'] - h['sku_text']  # QTY cell spans the SKU barcode and number rows
            rule(b)
            pdf.line(x, b + h['sku_text'], middle, b + h['sku_text'])
            item = items[slot] if slot < len(items) else None
            if item:
                draw_barcode(pdf, item['sku'], *left_cell, b + h['sku_text'], h['sku_bar'])
                cell_text(pdf, item['sku'], *left_cell, b, h['sku_text'], font['sku'])
                cell_text(pdf, item['qty'], *right_cell, b, y - b, font['qty'])
            y = b
            b = y - h['combined']
            rule(b, split=False)
            if item:  # template C11: ASN (14 digits) + SKU + QTY
                draw_barcode(pdf, job['asn'] + item['sku'] + str(item['qty']), *full, b, h['combined'])
            y = b
        b = y - h['ctn']
        shade(*left_cell, b, h['ctn'])
        rule(b)
        cell_text(pdf, 'CTN NUMBER / NO OF CTNS', *left_cell, b, h['ctn'], font['ctn_title'])
        cell_text(pdf, f"{carton['number']} of {len(job['cartons'])}", *right_cell, b, h['ctn'], font['ctn'])
        y = b
        b = y - h['skus_bar'] - h['skus_text']  # title spans the barcode and number rows
        shade(*left_cell, b, y - b)
        rule(b)
        cell_text(pdf, 'NO OF SKUS IN BOX', *left_cell, b, y - b, font['skus_title'])
        draw_barcode(pdf, len(items), *right_cell, b + h['skus_text'], h['skus_bar'])
        cell_text(pdf, len(items), *right_cell, b, h['skus_text'], font['skus'])
        pdf.rect(x, b, right - x, top - b)
        pdf.showPage()
    pdf.save()


def write_packing_pdf(job, path):
    pdf = canvas.Canvas(str(path), pagesize=A4)
    page = 0

    def header():
        nonlocal page
        page += 1
        for y, text in [(808, 'ASOS PACKING LIST'), (786, f"Supplier: {job['supplier']}"),
                        (767, f"ASOS PO: {job['po']}    ASN: {job['asn']}"),
                        (748, f"Destination: {job['destination']}    Colour: {job['colour']}"),
                        (729, f"Goods: {job['description']}"),
                        (710, f"Packaging: {len(job['cartons'])} CARTONS    Carton range: 1-{len(job['cartons'])}")]:
            draw_text(pdf, text, 28, y, A4[0] - 56, 12)
        for x, text in [(28, 'ASOS CTN'), (102, 'SOURCE CTN'), (200, 'SKU'), (320, 'SIZE'),
                        (395, 'QTY'), (460, 'MIXED SKU')]:
            draw_text(pdf, text, x, 680, 100, 10)
        pdf.line(28, 674, A4[0] - 28, 674)
        draw_text(pdf, f"Page {page}", 28, 24, 100, 9)
        return 653

    y, total = header(), 0
    for carton in job['cartons']:
        for item in carton['items']:
            if y < 65:
                pdf.showPage()
                y = header()
            for x, value in [(28, carton['number']), (102, carton['source']), (200, item['sku']),
                             (320, item['size']), (395, item['qty']),
                             (460, 'YES' if len(carton['items']) > 1 else 'NO')]:
                draw_text(pdf, value, x, y, 110, 11)
            total += item['qty']
            y -= 21
    if y < 65:
        pdf.showPage()
        y = header()
    draw_text(pdf, f"TOTAL: {total} UNITS / {len(job['cartons'])} CARTONS", 28, y - 12, 530, 13)
    pdf.save()


def export_job(job, out_dir, colour, shape='block'):
    if colour not in COLOURS or (shape not in MARK_SHAPES and COLOURS[colour] is not None):
        raise ValueError('请选择 ASN 颜色/形状')
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    target = out_dir / (job['asn'] + datetime.now().strftime('_%Y%m%d_%H%M%S_%f'))
    with tempfile.TemporaryDirectory(prefix='.asos_', dir=out_dir) as temp:
        folder = Path(temp) / 'output'
        folder.mkdir()
        write_labels(job, folder / 'CartonLabels.pdf', colour, shape)
        write_packing_pdf(job, folder / (job['asn'] + '.pdf'))
        base_target = target
        suffix = 0
        while True:
            try:
                folder.rename(target)
                break
            except FileExistsError:
                suffix += 1
                target = base_target.with_name(f'{base_target.name}_{suffix}')
    return target


def build_into(parent):
    controls = ttk.Frame(parent, padding=12)
    controls.pack(fill='x')
    controls.columnconfigure(1, weight=1)
    values = {key: tk.StringVar() for key in ('po', 'packing', 'supplier', 'description', 'out')}
    values['description'].set('Clothing')
    values['supplier'].set(load_cfg().get('supplier', ''))
    files = {'po': [], 'packing': []}

    patterns = {'po': '*.pdf', 'packing': '*.xlsx'}

    def add_files(key, paths):
        files[key] = sorted(set(files[key]) | {str(Path(p)) for p in paths})
        values[key].set(f"{len(files[key])} 个文件：" + '；'.join(Path(f).name for f in files[key]))

    def browse(key, folder=False):
        if key == 'out':
            result = filedialog.askdirectory(parent=parent)
            if result:
                values['out'].set(result)
        elif folder:
            result = filedialog.askdirectory(parent=parent)
            if result:
                add_files(key, Path(result).glob(patterns[key]))
        else:
            add_files(key, filedialog.askopenfilenames(parent=parent, filetypes=[
                ('ASOS PO PDF' if key == 'po' else '装箱资料表', patterns[key])]))

    def clear(key):
        files[key] = []
        values[key].set('')

    for row, (key, title) in enumerate([('po', 'ASOS PO（PDF）'), ('packing', '装箱资料表（xlsx）'),
                                        ('supplier', 'ASOS 供应商英文名'), ('description', '货物英文描述'),
                                        ('out', '输出文件夹')]):
        ttk.Label(controls, text=title).grid(row=row, column=0, sticky='e', padx=8, pady=4)
        state = 'readonly' if key in ('po', 'packing') else 'normal'
        ttk.Entry(controls, textvariable=values[key], state=state).grid(row=row, column=1, sticky='ew', pady=4)
        if key in ('po', 'packing', 'out'):
            actions = ttk.Frame(controls)
            actions.grid(row=row, column=2, sticky='w', padx=6)
            ttk.Button(actions, text='浏览…' if key == 'out' else '加文件…',
                       command=lambda k=key: browse(k)).pack(side='left')
            if key != 'out':
                ttk.Button(actions, text='加文件夹…', command=lambda k=key: browse(k, True)).pack(side='left', padx=4)
                ttk.Button(actions, text='清空', command=lambda k=key: clear(k)).pack(side='left')
    ttk.Label(controls, text='ASN 对照\n（每行：ASOS PO ID  ASN  合同号）').grid(
        row=5, column=0, sticky='ne', padx=8, pady=4)
    mapping_box = scrolledtext.ScrolledText(controls, height=6, font=('Consolas', 10))
    mapping_box.grid(row=5, column=1, sticky='ew', pady=4)
    ncs = tk.BooleanVar(value=False)
    ttk.Checkbutton(controls, text='NCS（Advanced Nominated Carrier Scheme，不用颜色标记）',
                    variable=ncs).grid(row=6, column=1, sticky='w', pady=4)
    ttk.Label(parent, text='按客人邮件逐行输入 ASOS PO ID、ASN 和合同号（COG PO #），例如：500038555238  42370000165328  PO159521。\n'
              '程序按合同号找装箱单（文件名需包含合同号，无关的装箱单自动忽略），再按颜色和目的仓匹配颜色组，并按 ASN 自动分配标记颜色。\n'
              '每个 ASN 箱号从 1 开始；原箱号保留在预览/箱单，外箱需同步改号。'
              '4 个 SKU 的箱使用 Extra SKUs 版式，超过 4 个会被阻止。按实际尺寸 100% 打印。',
              padding=(18, 5)).pack(anchor='w')
    buttons = ttk.Frame(parent, padding=8)
    buttons.pack()
    log = scrolledtext.ScrolledText(parent, font=('Consolas', 10), state='disabled', wrap='word')
    log.pack(fill='both', expand=True, padx=12, pady=(0, 12))

    def show(lines):
        log.configure(state='normal')
        log.delete('1.0', 'end')
        log.insert('end', '\n'.join(lines))
        log.configure(state='disabled')

    def show_plan(plan):
        jobs, errors = plan['jobs'], plan['errors']
        lines = [f"共 {len(jobs) + len(errors)} 个 PO：可生成 {len(jobs)}，失败 {len(errors)}", '']
        for job in jobs:
            units = sum(i['qty'] for c in job['cartons'] for i in c['items'])
            lines.append(f"[{mark_name(job)}] ASOS PO {job['po']} → ASN {job['asn']}  |  {job['destination']} "
                         f"{job['colour']}  |  {len(job['cartons'])} 箱 {units} 件")
            lines.append(f"    装箱单 {job['packing_file']}（内部 PO {job['internal_po']}  款号 {job['style']}）")
            if job['note']:
                lines.append(f"    ⚠ {job['note']}")
            for carton in job['cartons']:
                lines.append(f"    原箱 {carton['source']} → ASOS {carton['number']} of {len(job['cartons'])}"
                             f"  |  {len(carton['items'])} SKUs  |  {sum(i['qty'] for i in carton['items'])} 件")
                lines.extend(f"        {i['size']:6}  SKU {i['sku']}  × {i['qty']}" for i in carton['items'])
            lines.append('')
        if errors:
            lines += ['失败（不会生成）：'] + [f"    {e}" for e in errors] + ['']
        if plan['skipped']:
            lines += ['已跳过的装箱单：'] + [f"    {s}" for s in plan['skipped']]
        if plan['ignored']:
            lines.append(f"已忽略 {plan['ignored']} 个文件名不含所需合同号的 xlsx")
        if plan['other_pdfs']:
            lines.append(f"已忽略 {plan['other_pdfs']} 个非 ASOS PO 的 PDF")
        if plan['unused']:
            lines.append('ASN 对照中未选择 PO PDF 的 ASOS PO：' + '、'.join(plan['unused']))
        lines.append('请逐个核对 ASN 与客人邮件一致。')
        show(lines)

    def run(export=False):
        try:
            plan = plan_batch(files['po'], files['packing'], mapping_box.get('1.0', 'end'),
                              values['supplier'].get(), values['description'].get(), ncs.get())
            show_plan(plan)
            if not export:
                return
            if not plan['jobs']:
                raise ValueError('没有可生成的 ASN')
            if not values['out'].get().strip():
                raise ValueError('请选择输出文件夹')
            if plan['errors'] and not messagebox.askyesno(
                    '部分失败', f"{len(plan['errors'])} 个 PO 校验失败。是否先生成其余 {len(plan['jobs'])} 个？",
                    parent=parent):
                return
            targets = [export_job(job, values['out'].get(), job['mark'], job['shape']) for job in plan['jobs']]
            save_cfg(supplier=values['supplier'].get().strip())  # validated by make_job above
            messagebox.showinfo('生成完成', f"已生成 {len(targets)} 个 ASN 的标签及箱单：\n{values['out'].get()}"
                                '\n\n首次使用请先试印扫码。', parent=parent)
        except Exception as exc:
            show(['校验/生成失败：' + str(exc)])
            messagebox.showerror('无法生成', str(exc), parent=parent)

    ttk.Button(buttons, text='校验 / 预览', command=run).pack(side='left', padx=8)
    ttk.Button(buttons, text='批量生成标签 PDF + ASN 箱单 PDF', command=lambda: run(True)).pack(side='left', padx=8)


