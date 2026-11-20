# -*- coding: utf-8 -*-
"""BOM 物料分类汇总 — 从多个 Tech-Pack .XLS 文件提取 BOM 明细并分类。

每个款式对应一个 .XLS（旧版 BIFF/OLE2，必须用 xlrd 读取），
含名为「{PLM码} BOM」的工作表。程序抽取所有款式的 BOM 行，
按 11 个物料类别归类后合并写入一个 xlsx 输出文件。
"""

import json
import sys
import threading
import traceback
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

import generate_tabs as gt

if getattr(sys, 'frozen', False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).parent

OVERRIDES_FILE = APP_DIR / 'bom_category_overrides.json'
OUTPUT_FILENAME = 'BOM物料分类汇总.xlsx'
SHEET_NAME = 'BOM物料分类汇总'

# ── 类别定义（按顺序显示和排序，顺序即为输出排序依据）─────────────────
CATEGORIES = [
    'Microfiber',
    'Mesh',
    'Cups',
    'Elastic',
    'Strap',
    'Wire, Rings & Slides',
    'Hook & eye',
    'Wire casing',
    'Thread',
    'Tricot',
    'Fabric',
    '忽略/丢弃',
]
CATEGORY_ORDER = {c: i for i, c in enumerate(CATEGORIES)}

# ── 归类规则（按顺序匹配，命中即止）────────────────────────────────────
# 每条: (callable(desc_upper) -> bool, 类别名)
def _make_rules():
    def has(kw):
        return lambda d: kw in d

    def has_any(*kws):
        return lambda d: any(k in d for k in kws)

    def has_and_not(kw, excl):
        return lambda d: kw in d and excl not in d

    def has_any_and_not(kws, excl):
        return lambda d: any(k in d for k in kws) and excl not in d

    return [
        (has('HOOK'),                                   'Hook & eye'),
        (has_any('WIRE CASING', 'CASING'),              'Wire casing'),
        (has_any('METAL', 'WIRE', 'RING', 'SLIDE',
                 'ADJUSTER', 'BONING', 'UNDERWIRE'),    'Wire, Rings & Slides'),
        (has('STRAP'),                                  'Strap'),
        (has('ELASTIC'),                                'Elastic'),
        (has_any('CUP', 'FOAM'),                        'Cups'),
        (has('MICROFIB'),                               'Microfiber'),
        (has('MESH'),                                   'Mesh'),
        (has('TRICOT'),                                 'Tricot'),
        (has('THREAD'),                                 'Thread'),
        (has_any_and_not(('LACE', 'COTTON', 'JERSEY', 'ALLOVER'), 'TRIM'),
                                                        'Fabric'),
    ]


_RULES = _make_rules()


def categorize(desc: str):
    """Return category string or None if no rule matches."""
    d = desc.upper()
    for test, cat in _RULES:
        if test(d):
            return cat
    return None


# ── Override 持久化（同 qc_ship_overrides.json 套路）───────────────────
def load_overrides() -> dict:
    try:
        if OVERRIDES_FILE.exists():
            d = json.loads(OVERRIDES_FILE.read_text(encoding='utf-8'))
            if isinstance(d, dict):
                return d
    except Exception:
        pass
    return {}


def save_overrides(d: dict) -> None:
    try:
        OVERRIDES_FILE.write_text(
            json.dumps(d, ensure_ascii=False, indent=2), encoding='utf-8')
    except Exception:
        pass


def _override_key(desc: str) -> str:
    return desc.strip().upper()


# ── BOM 抽取（xlrd）────────────────────────────────────────────────────
def extract_bom_rows(folder: str) -> list[dict]:
    """Scan all .xls/.XLS files in folder and return list of {style, desc} dicts."""
    import xlrd  # deferred import; not available until xlrd is installed/included

    folder_path = Path(folder)
    rows = []
    for xls_path in sorted(folder_path.glob('*.xls')) + sorted(folder_path.glob('*.XLS')):
        try:
            wb = xlrd.open_workbook(str(xls_path), on_demand=True)
        except Exception:
            continue

        # Find the BOM sheet (ends with ' BOM')
        bom_sheet = None
        for name in wb.sheet_names():
            if name.strip().upper().endswith(' BOM') or name.strip().upper() == 'BOM':
                bom_sheet = wb.sheet_by_name(name)
                break
        if bom_sheet is None:
            continue

        # Style code from C2 (row 1, col 2)
        try:
            style = str(bom_sheet.cell_value(1, 2)).strip()
        except Exception:
            style = xls_path.stem

        # Remove trailing .0 if it's a numeric style code stored as float
        if style.endswith('.0') and style[:-2].isdigit():
            style = style[:-2]

        valid_types = {'Material', 'Component / Trim', 'Fabric'}
        for row_idx in range(bom_sheet.nrows):
            try:
                row_type = str(bom_sheet.cell_value(row_idx, 0)).strip()
            except Exception:
                continue
            if row_type not in valid_types:
                continue
            try:
                desc = str(bom_sheet.cell_value(row_idx, 3)).strip()
            except Exception:
                continue
            if not desc or desc == '.':
                continue
            rows.append({'style': style, 'desc': desc})

    return rows


# ── 未识别物料弹窗（Toplevel + Treeview + Combobox）────────────────────
def _ask_unknown(parent, unknown_descs: list, style_map: dict, overrides: dict) -> bool:
    """Show dialog to let user assign categories to unknown materials.

    Modifies `overrides` in-place and saves to disk.
    Returns True if user confirmed, False if cancelled.
    """
    result = {'confirmed': False}

    win = tk.Toplevel(parent)
    win.title('未识别物料 — 请指定类别')
    win.geometry('820x520')
    win.transient(parent)
    win.grab_set()

    tk.Label(win, fg='gray', justify='left', wraplength=800,
             text='以下物料无法自动归类，请为每条指定类别（或选「忽略/丢弃」跳过）。\n'
                  '双击「类别」列 → 从下拉选择，或直接输入自定义类别名称（回车确认）。'
                  '您的选择将被记忆，下次相同物料不再弹窗。'
             ).pack(side='top', fill='x', padx=8, pady=(8, 4))

    body = tk.Frame(win)
    body.pack(side='top', fill='both', expand=True)

    cols = ('desc', 'example_style', 'category')
    tree = ttk.Treeview(body, columns=cols, show='headings', height=16)
    tree.heading('desc', text='物料描述')
    tree.heading('example_style', text='出现款号（示例）')
    tree.heading('category', text='类别')
    tree.column('desc', width=340, anchor='w')
    tree.column('example_style', width=160, anchor='center')
    tree.column('category', width=200, anchor='w')

    vsb = ttk.Scrollbar(body, orient='vertical', command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    tree.pack(side='left', fill='both', expand=True, padx=(8, 0), pady=4)
    vsb.pack(side='left', fill='y', pady=4)

    # Build category list: standard ones + any custom ones already in overrides
    custom_in_overrides = sorted({v for v in overrides.values()
                                   if v and v not in CATEGORIES and v != '忽略/丢弃'})
    all_cats = CATEGORIES[:-1] + custom_in_overrides + ['忽略/丢弃']

    # choices dict: iid -> StringVar (category selection)
    choices = {}
    iid_to_desc = {}

    for desc in unknown_descs:
        examples = ', '.join(sorted(style_map.get(desc, set()))[:3])
        current = overrides.get(_override_key(desc), '')
        var = tk.StringVar(value=current or '忽略/丢弃')
        iid = tree.insert('', 'end', values=(desc, examples, var.get()))
        choices[iid] = (desc, var)
        iid_to_desc[iid] = desc

    # Inline Combobox editor
    combo_widget = {'w': None, 'iid': None}

    def _close_combo(_=None):
        if combo_widget['w'] is not None:
            iid = combo_widget['iid']
            if iid:
                desc, var = choices[iid]
                var.set(combo_widget['w'].get())
                tree.set(iid, 'category', var.get())
            combo_widget['w'].destroy()
            combo_widget['w'] = None
            combo_widget['iid'] = None

    def _open_combo(event):
        _close_combo()
        iid = tree.identify_row(event.y)
        if not iid or tree.identify_column(event.x) != '#3':
            return
        bbox = tree.bbox(iid, '#3')
        if not bbox:
            return
        x, y, w, h = bbox
        _, desc_var = choices[iid]
        cb = ttk.Combobox(tree, values=all_cats, state='normal',
                          width=w // 8)
        cb.set(desc_var.get())
        cb.place(x=x, y=y, width=w, height=h)
        cb.focus_set()

        def _commit(_=None):
            desc_var.set(cb.get())
            tree.set(iid, 'category', cb.get())
            cb.destroy()
            combo_widget['w'] = None
            combo_widget['iid'] = None

        cb.bind('<<ComboboxSelected>>', _commit)
        cb.bind('<Return>', _commit)
        cb.bind('<FocusOut>', _commit)
        cb.bind('<Escape>', lambda _: (cb.destroy(), None))
        combo_widget['w'] = cb
        combo_widget['iid'] = iid

    tree.bind('<Double-1>', _open_combo)

    def _save():
        _close_combo()
        for iid, (desc, var) in choices.items():
            overrides[_override_key(desc)] = var.get()
        save_overrides(overrides)
        result['confirmed'] = True
        win.destroy()

    def _cancel():
        result['confirmed'] = False
        win.destroy()

    btn_frame = tk.Frame(win)
    btn_frame.pack(side='bottom', fill='x', padx=8, pady=8)
    tk.Button(btn_frame, text='保存并继续', command=_save, width=14).pack(side='right', padx=4)
    tk.Button(btn_frame, text='取消', command=_cancel, width=8).pack(side='right', padx=4)
    tk.Label(btn_frame, text='双击「类别」列 → 选择或输入新类别名（回车/失焦确认）', fg='gray').pack(side='left')

    win.wait_window()
    return result['confirmed']


# ── 写出 xlsx ──────────────────────────────────────────────────────────
def write_output(classified: list[dict], out_path: Path) -> None:
    """Write classified BOM rows to xlsx.

    classified: list of {style, desc, category}
    """
    # Sort: category order first, then style
    def sort_key(r):
        return (CATEGORY_ORDER.get(r['category'], 999), r['style'], r['desc'])

    classified_sorted = sorted(classified, key=sort_key)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME

    headers = ['类别', '款号', '物料描述']
    bold = Font(bold=True)
    for col, hdr in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col, value=hdr)
        cell.font = bold
        cell.alignment = Alignment(horizontal='center', vertical='center')

    for row_idx, r in enumerate(classified_sorted, start=2):
        ws.cell(row=row_idx, column=1, value=r['category'])
        ws.cell(row=row_idx, column=2, value=r['style'])
        ws.cell(row=row_idx, column=3, value=r['desc'])

    # Auto column widths
    col_widths = [max(len(str(ws.cell(row=r, column=c).value or ''))
                      for r in range(1, ws.max_row + 1))
                  for c in range(1, 4)]
    for col, w in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(col)].width = min(w + 4, 60)

    out_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(str(out_path))


# ── 主生成流程 ─────────────────────────────────────────────────────────
def run_export(bom_src: str, bom_out: str, parent_widget, status_var, btn,
               on_need_dialog) -> None:
    """Run BOM extraction+classification in a background thread.

    on_need_dialog: called on main thread when unknown materials are found.
    """
    def worker():
        try:
            status_var.set('正在读取 BOM 文件…')
            rows = extract_bom_rows(bom_src)
            if not rows:
                status_var.set('未找到任何 BOM 数据，请检查文件夹和 .xls 文件。')
                btn.config(state='normal')
                return

            overrides = load_overrides()

            # Classify
            classified = []
            unknown_descs = []
            style_map: dict[str, set] = {}  # desc -> set of styles

            for r in rows:
                desc = r['desc']
                key = _override_key(desc)
                if key in overrides:
                    cat = overrides[key]
                else:
                    cat = categorize(desc)

                if cat is None:
                    if desc not in style_map:
                        style_map[desc] = set()
                        unknown_descs.append(desc)
                    style_map[desc].add(r['style'])
                else:
                    classified.append({**r, 'category': cat})

            if unknown_descs:
                status_var.set('发现未识别物料，请在弹窗中指定类别…')
                # Must call dialog on the main thread
                parent_widget.after(0, lambda: _handle_unknown(
                    rows, classified, unknown_descs, style_map, overrides,
                    bom_out, status_var, btn))
                return

            _finish(classified, bom_out, status_var, btn)

        except Exception:
            err = traceback.format_exc()
            status_var.set('错误：' + err.splitlines()[-1])
            btn.config(state='normal')

    def _handle_unknown(rows, classified, unknown_descs, style_map, overrides,
                        bom_out, status_var, btn):
        confirmed = _ask_unknown(parent_widget, unknown_descs, style_map, overrides)
        if not confirmed:
            status_var.set('已取消。')
            btn.config(state='normal')
            return
        # Re-classify with updated overrides (now saved to disk)
        classified2 = list(classified)
        for r in rows:
            desc = r['desc']
            key = _override_key(desc)
            if key in overrides:
                cat = overrides[key]
                if cat and cat != '忽略/丢弃':
                    classified2.append({**r, 'category': cat})
        _finish(classified2, bom_out, status_var, btn)

    def _finish(classified, bom_out, status_var, btn):
        try:
            # Drop 忽略/丢弃 rows
            classified = [r for r in classified if r.get('category') != '忽略/丢弃']

            out_path = Path(bom_out) / OUTPUT_FILENAME
            write_output(classified, out_path)

            # Summary by category
            cat_counts: dict[str, int] = {}
            for r in classified:
                cat_counts[r['category']] = cat_counts.get(r['category'], 0) + 1
            summary = '  '.join('%s×%d' % (c, n)
                                for c, n in sorted(cat_counts.items(),
                                                   key=lambda x: CATEGORY_ORDER.get(x[0], 999)))
            status_var.set('完成！共 %d 条物料已写出。%s' % (len(classified), summary))
            btn.config(state='normal')
        except Exception:
            err = traceback.format_exc()
            status_var.set('写出失败：' + err.splitlines()[-1])
            btn.config(state='normal')

    t = threading.Thread(target=worker, daemon=True)
    t.start()


# ── Tab UI ─────────────────────────────────────────────────────────────
def build_into(parent):
    cfg = gt.load_saved()
    p = {'padx': 8, 'pady': 4}
    parent.columnconfigure(1, weight=1)

    def _row(idx, label, key, default=''):
        tk.Label(parent, text=label).grid(row=idx, column=0, sticky='e', **p)
        var = tk.StringVar(value=cfg.get(key, '') or default)
        tk.Entry(parent, textvariable=var).grid(row=idx, column=1, sticky='ew', **p)
        return var

    def _pick_dir(var, cfgkey):
        path = filedialog.askdirectory()
        if path:
            var.set(path)
            gt.update_saved(**{cfgkey: path})

    src_var = _row(0, 'BOM 文件夹(.xls)：', 'bom_src')
    tk.Button(parent, text='文件夹…',
              command=lambda: _pick_dir(src_var, 'bom_src')
              ).grid(row=0, column=2, **p)

    out_var = _row(1, '输出文件夹：', 'bom_out',
                   default=cfg.get('output_waimao', ''))
    tk.Button(parent, text='文件夹…',
              command=lambda: _pick_dir(out_var, 'bom_out')
              ).grid(row=1, column=2, **p)

    tk.Label(parent,
             text='  说明：扫描 BOM 文件夹中所有 .xls/.XLS 文件（每款一个），提取 BOM 物料行，'
                  '按 11 个类别自动归类；无法归类的物料弹窗让您手动指定（选择将被记忆）。'
                  '输出文件：BOM物料分类汇总.xlsx，同名覆盖。',
             fg='gray', justify='left', wraplength=820).grid(
        row=2, column=0, columnspan=3, sticky='w', padx=8, pady=(2, 6))

    status_var = tk.StringVar(value='')
    tk.Label(parent, textvariable=status_var, fg='#005599', justify='left',
             wraplength=820).grid(row=4, column=0, columnspan=3, sticky='w', **p)

    generate_btn = tk.Button(parent, text='生成 BOM 分类汇总')
    generate_btn.grid(row=3, column=0, columnspan=2, pady=8)

    def _open_out():
        path = out_var.get().strip()
        if path and Path(path).is_dir():
            import subprocess
            subprocess.Popen(['explorer', Path(path)])

    tk.Button(parent, text='打开输出文件夹', command=_open_out).grid(
        row=3, column=2, **p)

    def _do_export():
        bom_src = src_var.get().strip()
        bom_out = out_var.get().strip()
        if not bom_src or not Path(bom_src).is_dir():
            messagebox.showwarning('提示', '请选择有效的 BOM 文件夹')
            return
        if not bom_out:
            messagebox.showwarning('提示', '请选择输出文件夹')
            return
        gt.update_saved(bom_src=bom_src, bom_out=bom_out)
        generate_btn.config(state='disabled')
        status_var.set('正在处理…')
        run_export(bom_src, bom_out, parent, status_var, generate_btn, None)

    generate_btn.config(command=_do_export)

    return _do_export
