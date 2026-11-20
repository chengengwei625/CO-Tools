"""
stamp_cartons.py
----------------
Reads packing-list CSVs from ./source/ and SSCC-label PDFs from ./labels/,
determines which carton each label page belongs to, then stamps the carton
number (BOX prefix stripped, e.g. "17") in the top-right corner of every page.

Stamped copies are written to ./labels_numbered/ (originals untouched).
"""

import csv
import datetime
import io
import os
import re
import sys
from collections import defaultdict
from itertools import groupby
from pathlib import Path

import fitz  # PyMuPDF


# ---------------------------------------------------------------------------
# Paths (resolved relative to the script / frozen exe location)
# ---------------------------------------------------------------------------

def base_dir() -> Path:
    # Use current working directory so the exe can be placed beside source/ and labels/
    return Path.cwd()


# ---------------------------------------------------------------------------
# CSV loading
# ---------------------------------------------------------------------------

def load_csv_cartons(source_dir: Path) -> dict[str, list[tuple[str, tuple]]]:
    """Return {po_number: [(box_name, frozenset_of_sku_rows), ...]} preserving box order."""
    po_boxes: dict[str, dict[str, list]] = {}
    po_order: dict[str, list[str]] = {}

    for csv_path in sorted(source_dir.glob("*.csv")):
        with open(csv_path, encoding="utf-8-sig", newline="") as fh:
            reader = csv.DictReader(fh)
            for row in reader:
                po    = (row.get("PurchaseOrder Number") or "").strip()
                box   = (row.get("Carton Name") or "").strip()
                sku   = (row.get("SKU") or "").strip()
                # 跳过空行/汇总行（CSV 末尾常有空行；SKU 必须为 款号-色号-尺码）
                parts = sku.split("-")
                if not po or not box or len(parts) < 3:
                    continue
                style, colour, size = parts[0], parts[1], "-".join(parts[2:])
                try:
                    qty = int(float((row.get("Quantity") or "0").strip() or 0))
                except (TypeError, ValueError):
                    continue
                if po not in po_boxes:
                    po_boxes[po] = {}
                    po_order[po] = []
                if box not in po_boxes[po]:
                    po_boxes[po][box] = []
                    po_order[po].append(box)
                po_boxes[po][box].append((style, colour, size, qty))

    result: dict[str, list[tuple[str, tuple]]] = {}
    for po, names in po_order.items():
        result[po] = [(name, tuple(sorted(po_boxes[po][name]))) for name in names]
    return result


# ---------------------------------------------------------------------------
# PDF text parsing
# ---------------------------------------------------------------------------

def parse_page_content(text: str) -> tuple:
    """
    Return a sorted tuple of (style, colour, size, qty) for one label page.
    Handles two fitz-extracted layouts:
      Multi-SKU table: lines like  "6339951-02\n<qty>\n<size>"
      Single-SKU block: "STYLE CODE\n<s>\nCOLOUR CODE\n<c>\nSIZE\n<z>\nQUANTITY\n<q>"
    """
    rows = []

    # Multi-SKU table: "6339951-02\n6\n14C"  repeated
    # Split on newlines and scan triplets (style-colour, qty, size)
    lines = [l.strip() for l in text.splitlines() if l.strip()]
    i = 0
    while i < len(lines) - 2:
        m = re.fullmatch(r"(\d{7})-(\d{2})", lines[i])
        if m and re.fullmatch(r"\d+", lines[i + 1]) and re.fullmatch(r"\w+", lines[i + 2]):
            style, colour = m.group(1), m.group(2)
            qty = int(lines[i + 1])
            size = lines[i + 2]
            rows.append((style, colour, size, qty))
            i += 3
        else:
            i += 1

    # Single-SKU block layout (fitz): each field on its own line
    if not rows:
        s = re.search(r"STYLE CODE\n(\d+)", text)
        c = re.search(r"COLOUR CODE\n(\d+)", text)
        z = re.search(r"SIZE\n(\w+)", text)
        q = re.search(r"QUANTITY\n(\d+)", text)
        if s and c and z and q:
            rows.append((s.group(1), c.group(1), z.group(1), int(q.group(1))))

    return tuple(sorted(rows))


def read_all_label_pages(labels_dir: Path, log_fn=print) -> list[dict]:
    """
    Return a list of page records sorted by global SSCC sequence number.
    Each record: {file, page_index, sscc_seq, po, content, sscc_barcode, destination}
    """
    records = []
    for pdf_path in sorted(labels_dir.glob("*.pdf")):
        doc = fitz.open(str(pdf_path))
        for page_index, page in enumerate(doc):
            text = page.get_text()
            # fitz puts the value on the next line after the label
            po_m   = re.search(r"PURCHASE ORDER\s*\n\s*(\S+)", text)
            sscc_m = re.search(r"SSCC-18\s*\n\s*(\d+) of (\d+)", text)
            if not po_m or not sscc_m:
                log_fn(f"  WARNING: could not parse PO/SSCC from {pdf_path.name} page {page_index+1}")
                continue
            # SSCC barcode value: digits after (00), formatted as "0039359198 XXXXXXXXXX"
            bc_m   = re.search(r"\(00\)(\d+)", text)
            dest_m = re.search(r"DESTINATION\s*\n\s*(.+)", text)
            raw_bc = bc_m.group(1) if bc_m else ""
            # Full SSCC = "00" + 18 raw digits = 20 digits; drop last (check digit) → 19 digits
            # Format as first-10 + space + next-9
            full_bc = "00" + raw_bc
            sscc_barcode = (full_bc[:10] + " " + full_bc[10:19]) if len(full_bc) >= 19 else full_bc
            records.append({
                "file":         pdf_path,
                "page_index":   page_index,
                "sscc_seq":     int(sscc_m.group(1)),
                "sscc_total":   int(sscc_m.group(2)),
                "po":           po_m.group(1).strip(),
                "content":      parse_page_content(text),
                "sscc_barcode": sscc_barcode,
                "destination":  dest_m.group(1).strip() if dest_m else "",
            })
        doc.close()

    records.sort(key=lambda r: r["sscc_seq"])
    return records


def validate_labels_vs_cartons(records, csv_cartons) -> dict:
    """处理前预校验：标签 PDF 与装箱单(CSV)的箱数是否一致。

    返回 {'missing_csv': [po,...],
          'count_mismatch': [(po, 标签页数, 装箱单箱数), ...],
          'total_mismatch': [(pdf文件名, 标签标示总数, 实际页数), ...]}。
    """
    # 各 PO 在标签 PDF 中的页数（按出现顺序；dict 保序）
    pdf_pages: dict = {}
    for r in records:
        pdf_pages[r["po"]] = pdf_pages.get(r["po"], 0) + 1

    missing_csv = [po for po in pdf_pages if po not in csv_cartons]
    count_mismatch = [
        (po, npages, len(csv_cartons[po]))
        for po, npages in pdf_pages.items()
        if po in csv_cartons and npages != len(csv_cartons[po])
    ]

    # 每个 PDF 文件内，标签自报的「of N」总数应等于该文件实际页数
    by_file: dict = defaultdict(list)
    for r in records:
        by_file[r["file"]].append(r)
    total_mismatch = []
    for f, recs in by_file.items():
        npages = len(recs)
        for n in sorted({r.get("sscc_total") for r in recs if r.get("sscc_total")}):
            if n != npages:
                total_mismatch.append((Path(f).name, n, npages))

    return {
        "missing_csv": missing_csv,
        "count_mismatch": count_mismatch,
        "total_mismatch": total_mismatch,
    }


# ---------------------------------------------------------------------------
# Mapping pages → carton names
# ---------------------------------------------------------------------------

def assign_carton_numbers(
    records: list[dict],
    csv_cartons: dict[str, list[tuple[str, tuple]]],
    log_fn=print,
) -> list[dict]:
    """
    For each page record, assign 'box_name' by grouping pages into PO blocks and
    reversing the CSV box list order (first page of a block = highest BOX number).
    Falls back to content-matching if the positional assignment disagrees.
    Returns the records list with 'box_name' and 'ctn_number' added.
    """
    warnings = []

    # Group globally-sorted records into contiguous PO runs
    for po, group_iter in groupby(records, key=lambda r: r["po"]):
        group = list(group_iter)
        if po not in csv_cartons:
            msg = f"  WARNING: PO '{po}' found in PDFs but has no matching CSV."
            log_fn(msg); warnings.append(msg)
            for rec in group:
                rec["box_name"] = "?"
                rec["ctn_number"] = "?"
            continue

        boxes = csv_cartons[po]          # ordered list [(name, content_tuple), ...]
        boxes_rev = list(reversed(boxes))  # reverse: first page → last box

        if len(group) != len(boxes):
            log_fn(f"  WARNING: PO {po} has {len(group)} pages but {len(boxes)} CSV boxes — counts differ!")

        content_index: dict[tuple, list[str]] = defaultdict(list)
        for name, content in boxes:
            content_index[content].append(name)
        content_pools: dict[tuple, list[str]] = {k: list(v) for k, v in content_index.items()}

        for i, rec in enumerate(group):
            # Primary: positional reverse assignment
            if i < len(boxes_rev):
                expected_name, expected_content = boxes_rev[i]
                if rec["content"] == expected_content:
                    rec["box_name"] = expected_name
                else:
                    # Fallback: content match
                    pool = content_pools.get(rec["content"], [])
                    if pool:
                        rec["box_name"] = pool.pop(0)
                        msg = (f"  WARNING: SSCC seq {rec['sscc_seq']} (PO {po}) positional mismatch "
                               f"— positional says {expected_name}, content match gives {rec['box_name']}")
                        log_fn(msg); warnings.append(msg)
                    else:
                        rec["box_name"] = "?"
                        msg = (f"  WARNING: SSCC seq {rec['sscc_seq']} (PO {po}) — "
                               f"could not match page to any box (content={rec['content']})")
                        log_fn(msg); warnings.append(msg)
            else:
                rec["box_name"] = "?"
                msg = f"  WARNING: SSCC seq {rec['sscc_seq']} (PO {po}) — page index {i} exceeds CSV box count {len(boxes)}"
                print(msg); warnings.append(msg)

            # Strip leading "BOX" prefix and append /total, e.g. "17/17"
            name = rec["box_name"]
            if name != "?":
                n = re.sub(r"^BOX", "", name, flags=re.IGNORECASE)
                rec["ctn_number"] = f"{n}/{len(boxes)}"
            else:
                rec["ctn_number"] = "?"

    return records, warnings


# ---------------------------------------------------------------------------
# Stamping
# ---------------------------------------------------------------------------

FONT_SIZE = 36          # base font size in points
MARGIN_TOP = 10         # pt from top edge
MARGIN_RIGHT = 10       # pt from right edge
MAX_NUM_WIDTH_FRAC = 0.35   # number must fit within this fraction of page width


def stamp_pdfs(records: list[dict], out_dir: Path, log_fn=print) -> None:
    """
    For each source PDF, open it, stamp all assigned carton numbers, and save
    to out_dir preserving the original filename.
    """
    out_dir.mkdir(parents=True, exist_ok=True)

    # Group records by source file
    file_records: dict[Path, list[dict]] = defaultdict(list)
    for rec in records:
        file_records[rec["file"]].append(rec)

    for pdf_path, recs in sorted(file_records.items()):
        doc = fitz.open(str(pdf_path))

        for rec in recs:
            if rec["ctn_number"] == "?":
                continue
            page = doc[rec["page_index"]]
            pw   = page.rect.width
            ph   = page.rect.height
            label = rec["ctn_number"]

            # Auto-scale font so 3-digit numbers still fit neatly
            font_size = FONT_SIZE
            # Helvetica-Bold character is roughly 0.6 × font_size wide
            estimated_w = len(label) * 0.6 * font_size
            max_w = pw * MAX_NUM_WIDTH_FRAC
            if estimated_w > max_w:
                font_size = max_w / (len(label) * 0.6)

            # Measure actual text width with fitz
            tw = fitz.get_text_length(label, fontname="helv", fontsize=font_size)
            x = pw - tw - MARGIN_RIGHT
            y = MARGIN_TOP + font_size       # fitz y is baseline

            page.insert_text(
                fitz.Point(x, y),
                label,
                fontname="helv",
                fontsize=font_size,
                color=(0, 0, 0),
            )

        out_path = out_dir / pdf_path.name
        out_path.write_bytes(doc.tobytes(garbage=4, deflate=True))
        doc.close()
        log_fn(f"  Saved: {out_path.name}")


# ---------------------------------------------------------------------------
# 装箱单 xlsx data loader
# ---------------------------------------------------------------------------

def load_packing_list_data(source_dir: Path, log_fn=None) -> dict:
    """
    Scan source_dir for 装箱单 xlsx files and extract port + colour info.

    Returns:
        {po_base: {'port': str, 'colour_map': {code_str: combined_label}}}

    - po_base    : e.g. 'PO147216'  (prefix before first '-' or '_')
    - port       : combined port string from the 装箱单, e.g.
                   'LOS ANGELES 洛杉矶 1000'  (English + Chinese + number)
    - colour_map : e.g. {'01': '黑色/BLACK/01', '02': '棕色/MAPLE SUGAR/02'}
                   keyed by 色号 code so it matches the 2-digit code in CSVs.
    """
    def _log(msg):
        if log_fn:
            log_fn(msg)

    try:
        import openpyxl
    except ImportError:
        _log("  警告: openpyxl 未安装，无法读取装箱单")
        return {}

    result: dict = {}
    xlsx_files = sorted(source_dir.glob("*.xlsx"))

    if not xlsx_files:
        _log("  (未找到装箱单 xlsx 文件，颜色/港口信息将使用默认值)")
        return {}

    for xlsx_path in xlsx_files:
        # Skip Excel lock files (temp files created while the file is open)
        if xlsx_path.name.startswith("~$"):
            continue

        _log(f"  读取装箱单: {xlsx_path.name}")
        try:
            wb = openpyxl.load_workbook(str(xlsx_path), data_only=True)
        except Exception as e:
            _log(f"    警告: 无法打开文件 — {e}")
            continue

        # PO from the PL filename (e.g. 'PO151715 6340032 … PL.xlsx' → PO151715) —
        # fallback for newer templates whose 合同号 cell moved / is blank.
        fn_m = re.match(r"(PO\d+|\d{5,})", xlsx_path.name.upper())
        file_po = fn_m.group(0) if fn_m else ""

        sheets_matched = 0
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            # Identify 装箱单 sheet: A3 must be "合同号:".  The PO/order value
            # moves between template versions (E3 old, D3 new) — scan row 3 for
            # the first PO-looking cell ("PO147216-1" / "771804" styles).
            a3 = ws["A3"].value
            if not a3 or str(a3).strip() != "合同号:":
                continue

            po_str = ""
            for c_idx in range(2, 12):                       # B3..K3
                v = str(ws.cell(3, c_idx).value or "").strip()
                if v and re.match(r"(PO\d+|\d{5,})", v.upper()):
                    po_str = v
                    break
            if not po_str:
                po_str = file_po                             # 合同号格为空 → 用文件名
            if not po_str:
                continue

            # Base PO: everything before first '-' or '_' (e.g. PO147216-1→PO147216, 771804→771804)
            po_base = re.split(r"[-_]", po_str)[0]
            sheets_matched += 1

            # ── Find colour summary section (row with A = '主布颜色') ──────
            # Strip all whitespace variants when comparing
            zhbu_row = None
            for r_idx in range(1, min(ws.max_row + 1, 200)):
                cell_val = ws.cell(r_idx, 1).value
                if cell_val and str(cell_val).strip() == "主布颜色":
                    zhbu_row = r_idx
                    break

            if not zhbu_row:
                _log(f"    警告: 工作表 '{sheet_name}' 中未找到 '主布颜色' 行")

            # ── Port info ──────────────────────────────────────────────────
            # 港口(英文+号)在第 2 行，列因模板版本而异（G2 旧版 / H2 新版等）：
            # 扫描 B2..K2 取第一个「含英文字母、无换行、不超长」的格（跳过 J/K 的长备注）。
            # 中文港口名取同列第 3 行（G2→G3、H2→H3），为空再退回 G3/F3。
            port_col_str = ""
            port_col_idx = 0
            for c_idx in range(2, 12):                       # B2..K2
                v = str(ws.cell(2, c_idx).value or "").strip()
                if (v and "\n" not in v and len(v) <= 40
                        and re.search(r"[A-Za-z]", v)):
                    port_col_str, port_col_idx = v, c_idx
                    break
            port_cn_cell = ""
            if port_col_idx:
                port_cn_cell = str(ws.cell(3, port_col_idx).value or "").strip()
            if not port_cn_cell:
                port_cn_cell = (str(ws["G3"].value or "").strip()
                                or str(ws["F3"].value or "").strip())
            a_port = ""
            if zhbu_row and zhbu_row > 1:
                ap = ws.cell(zhbu_row - 1, 1).value
                if ap:
                    a_port = str(ap).strip()
            port_str = a_port or port_col_str

            # 港口号：port_col_str 中的数字（"SANTOS - 1500" → "1500"）
            num_m = re.search(r"\d+", port_col_str)
            port_no = num_m.group(0) if num_m else ""
            # 中文港口名：优先取港口同列第 3 行整格（如 'ASOS EU 欧洲' 保留英文前缀），
            # 该格无中文时再从 port_str 提取中文
            if re.search(r"[一-鿿]", port_cn_cell):
                port_cn = port_cn_cell
            else:
                cn_m = re.search(r"[一-鿿]+", port_str)
                port_cn = (cn_m.group(0) if cn_m else "") or port_cn_cell
            # 英文匹配词（用于按标签目的地城市匹配，如 "SANTOS"/"BRAZIL"）
            tokens = set(re.findall(r"[A-Za-z]+", (port_col_str + " " + a_port).upper()))
            port_entry = {"cn": port_cn, "no": port_no, "name": port_str,
                          "col": port_col_str, "tokens": tokens}

            # ── Colour map from rows below the 主布颜色 header ───────────
            # 客色号/色号 列因模板版本而异（旧版 D/E，新版 B/E）：按表头定位
            colour_map: dict[str, str] = {}
            if zhbu_row:
                en_col, code_col = 4, 5
                for c_idx in range(2, 12):
                    h = str(ws.cell(zhbu_row, c_idx).value or "").strip()
                    if re.fullmatch(r"客色[号號]", h):
                        en_col = c_idx
                    elif re.fullmatch(r"色[号號]", h):
                        code_col = c_idx
                for cr in range(zhbu_row + 1, zhbu_row + 20):
                    zh   = ws.cell(cr, 1).value        # 主布颜色 (Chinese)
                    en   = ws.cell(cr, en_col).value   # 客色号  (English)
                    code = ws.cell(cr, code_col).value # 色号    (code like '01')
                    if not zh and not en:
                        break
                    if zh and en:
                        zh_s   = str(zh).strip()
                        en_s   = str(en).strip()
                        code_s = str(code).strip() if code else ""
                        label  = f"{zh_s}/{en_s}/{code_s}" if code_s else f"{zh_s}/{en_s}"
                        if code_s:
                            colour_map[code_s] = label

            if colour_map:
                _log(f"    {po_base}: 港口={port_str}, 颜色={list(colour_map.values())}")
            else:
                _log(f"    警告: 工作表 '{sheet_name}' 未提取到颜色数据")

            # Merge into result. 一个 PO 可发往多个港口（每个目的地一个工作表），
            # 全部收进 port_list，导出时按标签目的地城市匹配出正确的中文港口。
            if po_base not in result:
                result[po_base] = {"port": port_str, "port_cn": port_cn,
                                   "port_no": port_no, "colour_map": colour_map,
                                   "port_list": []}
            else:
                result[po_base]["colour_map"].update(colour_map)
            if tokens or port_no or port_cn:
                result[po_base]["port_list"].append(port_entry)

        if sheets_matched == 0:
            _log(f"    (文件中未发现有效装箱单工作表，已跳过)")

    return result


# ---------------------------------------------------------------------------
# XLSX export
# ---------------------------------------------------------------------------

def _dominant_colour(content: tuple) -> str:
    """Return the most common colour code across all SKU rows in a carton."""
    if not content:
        return ""
    from collections import Counter
    return Counter(colour for (_s, colour, _z, _q) in content).most_common(1)[0][0]


def export_xlsx(
    records: list[dict],
    csv_cartons: dict[str, list[tuple[str, tuple]]],
    out_dir: Path,
    log_fn=print,
    source_dir: Path | None = None,
) -> None:
    """
    Write one xlsx file per PO to out_dir.
    Format mirrors the 装箱单 template:
      Row 1: 目的地  | <port from 装箱单, or destination from label>
      Row 2: 合同号  | <PO with _suffix stripped, e.g. 'PO147216'>
      Row 3: 款式号  | <style(s)>
      Row 4: 颜色 | 箱号 | SSCC条码  (grey header)
      Rows…: combined colour (主布颜色/客色號/色号, first row of group only) | box# | SSCC
      Last:  日期    | <today's date>
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    out_dir.mkdir(parents=True, exist_ok=True)

    # Load 装箱单 data for enriched port & colour labels
    packing_data = load_packing_list_data(source_dir, log_fn=log_fn) if source_dir else {}

    today_str = datetime.date.today().strftime("%Y-%m-%d")

    # Group records by PO, preserving encounter order
    po_groups: dict[str, list[dict]] = {}
    for rec in records:
        po_groups.setdefault(rec["po"], []).append(rec)

    thin = Side(style="thin")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for po, recs in po_groups.items():
        # Base PO for 合同号 display: strip '_' and everything after
        po_display = re.sub(r"_.*$", "", po)

        # Base PO for 装箱单 lookup: same as display (PO147216)
        pl = packing_data.get(po_display, {})
        colour_map = pl.get("colour_map", {})
        port_list = pl.get("port_list", [])

        # Match port: sticker port number → 装箱单 entry → Chinese name.
        # Step 1: exact port number match (preferred).
        # Step 2: if no exact match, token match against entries whose G2 was blank
        #         (port_no='') — those PLs omitted the number, so number match is
        #         impossible; token match is the only option. Port number is then
        #         taken from the sticker.
        label_dest = next((r["destination"] for r in recs if r.get("destination")), "")
        label_num_m = re.search(r"\d+", label_dest)
        label_port_no = label_num_m.group(0) if label_num_m else ""
        matched = None
        if label_port_no and port_list:
            matched = next((pe for pe in port_list if pe["no"] == label_port_no), None)
        if matched is None and port_list:
            # Step 2: token match（装箱单港口格无号码时按目的地城市英文词匹配）
            label_tokens = set(re.findall(r"[A-Za-z]+", label_dest.upper()))
            cands = [pe for pe in port_list if pe["tokens"] & label_tokens]
            if cands:
                matched = max(cands, key=lambda pe: len(pe["tokens"] & label_tokens))
        if matched is None and port_list:
            log_fn(
                f"  ⚠ PO {po_display}: 标签港口号 '{label_port_no}'"
                f"（来自 '{label_dest}'）未在装箱单中找到对应港口，"
                f"文件名/目的地将使用标签原文"
            )
        port_cn = matched["cn"] if matched else ""
        port_no = (matched["no"] or label_port_no) if matched else label_port_no

        # Sort by box number ascending (1 → N)
        def _box_int(r):
            ctn = r.get("ctn_number", "0/1")
            try:
                return int(ctn.split("/")[0])
            except ValueError:
                return 0
        recs_sorted = sorted(recs, key=_box_int)

        # Destination: 英文港口名+号码（G2/F2） + 中文港口名（G3/F3）→ "SANTOS - 1500 巴西"
        if matched:
            col = matched.get("col", "")   # e.g. "SANTOS - 1500"
            cn  = matched["cn"]            # e.g. "巴西"
            destination = (col + " " + cn).strip() if col else matched["name"]
        else:
            destination = label_dest
        styles = sorted({s for r in recs for (s, _c, _z, _q) in r.get("content", ())})
        style_str = " / ".join(styles)

        wb = Workbook()
        ws = wb.active
        ws.title = "Sheet1"

        # ── Header rows 1-3 ──────────────────────────────────────────────
        header_font = Font(name="Arial", bold=True, size=10)
        value_font  = Font(name="Arial", size=10)

        def hrow(label, value):
            ws.append([label, value])
            r = ws.max_row
            ws.cell(r, 1).font = header_font
            ws.cell(r, 2).font = value_font

        hrow("目的地", destination)
        hrow("合同号", po_display)
        hrow("款式号", style_str)

        # ── Column header row ────────────────────────────────────────────
        ws.append(["颜色", "箱号", "SSCC条码"])
        col_hdr_fill = PatternFill("solid", fgColor="D9D9D9")
        col_hdr_font = Font(name="Arial", bold=True, size=10)
        for cell in ws[ws.max_row]:
            if cell.value:
                cell.font      = col_hdr_font
                cell.fill      = col_hdr_fill
                cell.border    = border
                cell.alignment = Alignment(horizontal="center")

        # ── Data rows ────────────────────────────────────────────────────
        colour_ranges = []   # (label, first_excel_row, last_excel_row)
        prev_colour   = None
        colour_start  = None

        for rec in recs_sorted:
            raw_colour        = _dominant_colour(rec.get("content", ()))
            colour_label_full = colour_map.get(raw_colour, raw_colour)

            ctn   = rec.get("ctn_number", "?")
            box_n = int(ctn.split("/")[0]) if "/" in ctn else ctn
            sscc  = rec.get("sscc_barcode", "")

            ws.append(["", box_n, sscc])
            r = ws.max_row

            if colour_label_full != prev_colour:
                if prev_colour is not None:
                    colour_ranges.append((prev_colour, colour_start, r - 1))
                colour_start = r
                prev_colour  = colour_label_full

            for col in (2, 3):
                ws.cell(r, col).font   = Font(name="Arial", size=10)
                ws.cell(r, col).border = border
            ws.cell(r, 2).alignment = Alignment(horizontal="center")

        # Close last colour group
        if prev_colour is not None:
            colour_ranges.append((prev_colour, colour_start, ws.max_row))

        # Merge colour column cells per group and apply style
        for clabel, start_r, end_r in colour_ranges:
            if end_r > start_r:
                ws.merge_cells(f"A{start_r}:A{end_r}")
            cell = ws.cell(start_r, 1)
            cell.value     = clabel
            cell.font      = Font(name="Arial", size=10)
            cell.border    = border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # ── Date row (last row) ───────────────────────────────────────────
        ws.append(["日期", today_str])
        r = ws.max_row
        date_font = Font(name="Arial", bold=True, size=10)
        ws.cell(r, 1).font   = date_font
        ws.cell(r, 2).font   = Font(name="Arial", size=10)
        ws.cell(r, 1).border = border
        ws.cell(r, 2).border = border

        # ── Column widths ────────────────────────────────────────────────
        ws.column_dimensions["A"].width = 26
        ws.column_dimensions["B"].width = 8
        ws.column_dimensions["C"].width = 24

        # ── Save （文件名：PO - 中文港口名 港口号；无匹配时退回标签原文）──
        if matched:
            port_suffix = " ".join(x for x in (port_cn, port_no) if x).strip()
        else:
            port_suffix = label_dest.strip()
        stem = f"{po_display} - {port_suffix}" if port_suffix else po_display
        stem = re.sub(r'[\\/:*?"<>|]', "-", stem).strip()
        buf = io.BytesIO()
        wb.save(buf)
        out_path = out_dir / f"{stem}.xlsx"
        out_path.write_bytes(buf.getvalue())
        log_fn(f"  导出: {out_path.name}")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    base     = base_dir()
    src_dir  = base / "source"
    lbl_dir  = base / "labels"
    out_dir  = base / "labels_numbered"

    print("=== Stamp Carton Numbers onto SSCC Labels ===\n")

    # 1. Load CSV carton data
    print("Loading CSV packing data …")
    csv_cartons = load_csv_cartons(src_dir)
    for po, boxes in csv_cartons.items():
        print(f"  {po:20s}  {len(boxes):3d} boxes")
    print()

    # 2. Read label PDF pages
    print("Reading label PDF pages …")
    records = read_all_label_pages(lbl_dir)
    print(f"  {len(records)} pages read across {len(set(r['file'] for r in records))} PDF files\n")

    # 3. Assign carton numbers
    print("Assigning carton numbers …")
    records, warnings = assign_carton_numbers(records, csv_cartons)
    print()

    # 4. Stamp and save
    print("Stamping and saving PDFs …")
    stamp_pdfs(records, out_dir)
    print()

    # 5. Export xlsx
    print("Exporting SSCC xlsx files …")
    export_xlsx(records, csv_cartons, out_dir, source_dir=src_dir)
    print()

    # 6. Summary
    stamped = sum(1 for r in records if r.get("ctn_number", "?") != "?")
    print("=== Done ===")
    print(f"  Pages stamped : {stamped} / {len(records)}")
    print(f"  Warnings      : {len(warnings)}")
    if warnings:
        print("  See lines above for details.")
    print(f"  Output folder : {out_dir}")

    # Per-PO breakdown
    per_po: dict[str, int] = defaultdict(int)
    for r in records:
        if r.get("ctn_number", "?") != "?":
            per_po[r["po"]] += 1
    print("\n  Per-PO counts:")
    for po, count in per_po.items():
        print(f"    {po:20s}  {count}")


if __name__ == "__main__":
    main()
