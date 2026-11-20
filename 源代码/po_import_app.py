"""
PO → 订单明细 导入工具  v1.4
Changes:
  - Tabbed UI (文件 / 配置 / 预览)
  - 英文颜色 pre-filled from 客户色名, user-editable
  - Ctrl+V paste into colour table (TSV from Excel)
  - Template & port paths persisted across sessions
  - If same COLOUR DESCRIPTION appears with multiple STYLE COLOUR CODEs
    → one output file per style code
"""
import os, json
import re
from calendar import Calendar
from queue import Empty, Queue
from threading import Thread
from tempfile import NamedTemporaryFile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from copy import copy
from datetime import date, datetime
from collections import defaultdict

from po_pdf import load_po, load_fac_style_map
from erp_credentials import CredentialError, CredentialStore
from order_erp_client import ERP_URL, LIBRARIES, ErpError, OrderErpClient


COLOUR_IMPORT_FILL = PatternFill(fill_type='solid', fgColor='FFE2F0D9')
OUTPUT_FONT_NAME = '微软雅黑'
CO_CODE_PATTERN = re.compile(r'CO([0-9]+)', re.IGNORECASE)

# ERP 下拉选项和用于自动推荐的常见中英文颜色词。
ERP_COLOUR_TERMS = (
    ('白色(W)', '白|象牙|奶油|本白', 'off white|white|ivory|cream|ecru|snow'),
    ('黑色(B)', '黑|墨黑', 'black|onyx|ebony'),
    ('灰色(H)', '灰|炭灰', 'grey|gray|charcoal|slate|heather gray|heather grey'),
    ('红色(R)', '红|酒红|枣红|勃艮第', 'red|burgundy|maroon|scarlet|crimson|wine|cherry|ruby'),
    ('橙色(A)', '橙|橘|杏橙|珊瑚', 'orange|tangerine|pumpkin|rust|coral'),
    ('黄色(Y)', '黄|姜黄|芥末|柠檬', 'yellow|mustard|lemon'),
    ('绿色(G)', '绿|橄榄|鼠尾草', 'green|olive|khaki green|sage|mint|teal|emerald|lime'),
    ('蓝色(L)', '蓝|靛|藏青|宝蓝|粉蓝', 'blue|navy|indigo|denim|cobalt|turquoise|aqua'),
    ('紫色(P)', '紫|薰衣草|熏衣草|粉紫', 'purple|lilac|lavender|violet|plum|mauve'),
    ('粉红色(K)', '粉|玫红|桃红|桃粉', 'pink|fuchsia|magenta|blush|rose|peach'),
    ('肤色(F)', '肤|裸|米色|浅米|裸粉', 'nude|skin|beige|sand|stone'),
    ('咖啡色(C)', '咖啡|棕|褐|驼|卡其|巧克力|焦糖', 'brown|coffee|chocolate|tan|camel|khaki|taupe|mocha|caramel'),
    ('金色(J)', '金色|金黄|香槟金|玫瑰金|金', 'rose gold|gold|golden|champagne'),
    ('银色(S)', '银', 'silver'),
    ('再加工(Z)', '再加工', 'rework'),
    ('混合色(M)', '混合色', 'mixed color|mixed colour|multi color|multi colour|multicolor|multicolour'),
    ('花沙色(T)', '花沙色|花沙', ''),
    ('待定色(X)', '待定色|颜色待定', 'tbd|to be determined'),
)
ERP_COLOUR_CATEGORIES = tuple(category for category, _, _ in ERP_COLOUR_TERMS)
_ERP_COLOUR_ALIASES = {category: (chinese, english) for category, chinese, english in ERP_COLOUR_TERMS}

# 成品分类与物料分类使用不同的 ERP 代码，按成品下拉框顺序列出。
ERP_FINISHED_COLOUR_TERMS = (
    ('混合系1(A)', '混合系1|混合色1', 'mixed 1|mix 1'),
    ('蓝色系(B)', '蓝|靛|藏青|宝蓝|粉蓝', 'blue|navy|indigo|denim|cobalt'),
    ('青色系(C)', '青色|青绿|蓝绿|湖蓝', 'cyan|teal|turquoise|aqua'),
    ('咖啡色(D)', '咖啡|摩卡|巧克力', 'coffee|mocha|chocolate|espresso'),
    ('灰1系(E)', '灰1系|灰1色|灰色1', 'gray 1|grey 1'),
    ('肤色系(F)', *_ERP_COLOUR_ALIASES['肤色(F)']),
    ('绿色系(G)', '绿|橄榄|鼠尾草', 'green|olive|khaki green|sage|mint|emerald|lime'),
    ('灰色系(H)', '灰|银', 'grey|gray|charcoal|slate|silver'),
    ('灰2系(J)', '灰2系|灰2色|灰色2', 'gray 2|grey 2'),
    ('黑色系(K)', *_ERP_COLOUR_ALIASES['黑色(B)']),
    ('红1系(L)', '红1系|红1色|红色1', 'red 1'),
    ('黄1系(M)', '黄1系|黄1色|黄色1', 'yellow 1'),
    ('橙色系(N)', *_ERP_COLOUR_ALIASES['橙色(A)']),
    ('紫色系(P)', *_ERP_COLOUR_ALIASES['紫色(P)']),
    ('蓝1系(Q)', '蓝1系|蓝1色|蓝色1', 'blue 1'),
    ('红色系(R)', *_ERP_COLOUR_ALIASES['红色(R)']),
    ('紫1系(S)', '紫1系|紫1色|紫色1', 'purple 1'),
    ('白色系(W)', *_ERP_COLOUR_ALIASES['白色(W)']),
    ('黄色(Y)', '黄|芥末|柠檬|金色|金黄', 'yellow|mustard|lemon|gold|golden'),
    ('混合系(Z)', '混合系|混合色|多色|拼色',
     'mixed color|mixed colour|multi color|multi colour|multicolor|multicolour|assorted|mix'),
    ('粉色系(T)', *_ERP_COLOUR_ALIASES['粉红色(K)']),
    ('印花系(I)', '印花', 'print|printed'),
    ('棕色系(V)', '棕|褐|驼|卡其|焦糖', 'brown|tan|camel|khaki|taupe|caramel'),
)
ERP_FINISHED_COLOUR_CATEGORIES = tuple(category for category, _, _ in ERP_FINISHED_COLOUR_TERMS)
ERP_FINISHED_SPECIAL_CATEGORIES = {'混合系1(A)', '灰1系(E)', '灰2系(J)', '红1系(L)', '黄1系(M)',
                                   '蓝1系(Q)', '紫1系(S)', '混合系(Z)', '印花系(I)'}


def _erp_colour_pattern(chinese, english):
    patterns = sorted(chinese.split('|'), key=len, reverse=True)
    if english:
        english_terms = '|'.join(sorted(english.split('|'), key=len, reverse=True))
        patterns.append(rf'(?<![a-z])(?:{english_terms})(?![a-z])')
    return re.compile('|'.join(patterns), re.IGNORECASE)


ERP_COLOUR_PATTERNS = tuple(
    (category, _erp_colour_pattern(chinese, english))
    for category, chinese, english in ERP_COLOUR_TERMS
)


ERP_FINISHED_COLOUR_PATTERNS = tuple(
    (category, _erp_colour_pattern(chinese, english))
    for category, chinese, english in ERP_FINISHED_COLOUR_TERMS
)


def _match_erp_colour_category(name, patterns, priority_categories=()):
    matches = []
    for category, pattern in patterns:
        match = pattern.search(name)
        if match:
            matches.append((category not in priority_categories, match.start(), -len(match.group()), category))
    return min(matches)[-1] if matches else ''


def _suggest_erp_colour_category(colour_name, english_name, *, finished=False):
    patterns = ERP_FINISHED_COLOUR_PATTERNS if finished else ERP_COLOUR_PATTERNS
    priority_categories = ERP_FINISHED_SPECIAL_CATEGORIES if finished else ()
    for name in (colour_name, english_name):
        name = ' '.join(name.replace('-', ' ').split())
        category = _match_erp_colour_category(name, patterns, priority_categories)
        if finished and category and category not in priority_categories:
            parts = re.split(r'[/+&／＋＆]', name)
            if len(parts) > 1:
                component_categories = {_match_erp_colour_category(part, patterns) for part in parts}
                component_categories.discard('')
                if len(component_categories) > 1:
                    return '混合系(Z)'
        if category:
            return category
    return ''


def _erp_colour_list_rows(mapping_rows, initial_max_co=None, existing_customer_codes=frozenset()):
    rows = {}
    for code_column in (0, 2):
        for mapping in mapping_rows:
            code = mapping[code_column].strip()
            if code:
                rows[(code, mapping[3].strip(), mapping[1].strip())] = None

    new_rows = []
    for row in rows:
        code = row[0]
        if code.startswith('6') and code.casefold() in existing_customer_codes:
            continue
        if initial_max_co is not None:
            match = CO_CODE_PATTERN.fullmatch(code)
            if match and int(match.group(1)) <= initial_max_co:
                continue
        new_rows.append(row)
    return tuple(new_rows)


def _set_output_font(cell):
    font = copy(cell.font)
    font.name = OUTPUT_FONT_NAME
    font.scheme = None
    cell.font = font


def _clip_list(items, n=15):
    """汇总弹窗里列清单，超长截断。"""
    body = '\n'.join(f'  {x}' for x in items[:n])
    return body + (f'\n  … 等 {len(items)} 个' if len(items) > n else '')

FREIGHT_MAP = {
    'SEA': '船运-一般船运',
    'AIR': '空运',
    'LAND': '陆运',
    'EXPRESS': '快递',
}

CONFIG_FILE = os.path.join(os.path.expanduser('~'), '.po_import_tool.json')

# Columns that are never sizes — used to locate the size zone in the template
KNOWN_FIXED_COLS = {
    '颜色编号', '颜色名称', '英文颜色', '走货方式', '流水号', '目的港', '港口代码',
    '实际出厂（年-月-日）', '客户船期（年-月-日）', '杯码/内长', '价格',
    '客户色号', '客户色名', 'Payment Terms', 'CMT Price/Pc', 'Amount(USD)',
    '比例', '数量合计', '客款号', '交期1', '厂款号', '厂款号2', '工厂出厂期',
    '英文颜色拆分', '厂款号拆分',
}
POST_SIZE_COLS = {'比例', '数量合计', '客款号', '交期1', '英文颜色拆分', '厂款号拆分'}

# ── Size sort key ─────────────────────────────────────────────────────────
# Rank for standard letter-based garment sizes (XS < S < M < L < XL …)
_STD_SIZE_RANK = {
    'XXXXS': -6, '4XS': -6,
    'XXXS':  -5, '3XS': -5,
    'XXS':   -4, '2XS': -4,
    'XS':    -3,
    'S':      0,
    'M':      2,
    'L':      4,
    'XL':     6,
    'XXL':    7, '2XL': 7,
    'XXXL':   8, '3XL': 8,
    'XXXXL':  9, '4XL': 9,
    'XXXXXL':10, '5XL':10,
    '6XL':   11,
}

def _size_sort_key(s):
    """
    Sort key so sizes go small → large:
      • Standard letter sizes (XS < S < M < L < XL < XXL …)
      • Numeric-prefix sizes  (10A < 10B < 12A … or 36 < 37 …)
      • Unknown: alphabetical fallback
    """
    u = str(s).upper().strip()
    if u in _STD_SIZE_RANK:
        return (0, _STD_SIZE_RANK[u], '')
    digits  = ''.join(c for c in u if c.isdigit())
    letters = ''.join(c for c in u if not c.isdigit())
    if digits:
        return (1, int(digits), letters)
    return (2, 0, u)


# ── Inline-editable Treeview ──────────────────────────────────────────────
class InlineEditTreeview(ttk.Treeview):
    def __init__(self, parent, editable_cols, paste_table=None, batch_columns=(), **kw):
        super().__init__(parent, **kw)
        self.editable_cols = editable_cols
        self.paste_table = paste_table
        self.batch_columns = set(batch_columns)
        self._batch_column = None
        self._editor = None
        self._batch_rows = {}
        self._batch_click = False
        # Let Treeview select the row before the editor takes keyboard focus.
        self.bind('<ButtonRelease-1>', self._start_edit)
        if self.batch_columns:
            self.bind('<Control-Button-1>', self._toggle_batch_cell)
            self.bind('<Control-Double-Button-1>', self._toggle_batch_cell)
            self.bind('<Return>', self._start_batch_edit)
            self.bind('<Escape>', self._cancel_batch_selection)

    def _toggle_batch_cell(self, event):
        if self.identify_region(event.x, event.y) != 'cell':
            return
        column = self.column(self.identify_column(event.x), 'id')
        row = self.identify_row(event.y)
        if not row or column not in self.batch_columns:
            return
        if self._batch_column is not None and column != self._batch_column:
            self._cancel_batch_selection()
        if not self._batch_rows and self._editor is not None:
            _, targets, edit_column, _ = self._editor
            if edit_column == column:
                self._batch_rows[targets[0]] = None
        self.finish_edit()
        if row in self._batch_rows:
            del self._batch_rows[row]
        else:
            self._batch_rows[row] = None
        self._batch_column = column if self._batch_rows else None
        self.selection_set(tuple(self._batch_rows))
        self.focus(row)
        self.focus_set()
        self._batch_click = True
        return 'break'

    def _cancel_batch_selection(self, event=None):
        if self._batch_rows:
            self.finish_edit()
            self.selection_remove(tuple(self._batch_rows))
            self._batch_rows.clear()
            self._batch_column = None
            return 'break'

    def _start_batch_edit(self, event=None):
        for row in reversed(self._batch_rows):
            if self.bbox(row, self._batch_column):
                self._edit_cell(row, self._batch_column, tuple(self._batch_rows))
                return 'break'

    def _start_edit(self, event):
        if self._batch_click:
            self._batch_click = False
            return 'break'
        self.finish_edit()
        region = self.identify_region(event.x, event.y)
        row = self.identify_row(event.y) if region == 'cell' else ''
        display_column = self.identify_column(event.x) if row else ''
        column = self.column(display_column, 'id') if display_column else None
        if column == self._batch_column and row in self._batch_rows:
            self.selection_set(tuple(self._batch_rows))
            self._edit_cell(row, column, tuple(self._batch_rows))
            return
        if self._batch_rows:
            self._cancel_batch_selection()
            if row:
                self.selection_add(row)
        if row and int(display_column[1:]) - 1 in self.editable_cols:
            self._edit_cell(row, column)

    def _edit_cell(self, row, column, batch_rows=()):
        bbox = self.bbox(row, column)
        if not bbox:
            return
        x, y, w, h = bbox
        entry = tk.Entry(self, font=('Arial', 9))
        entry.place(x=x, y=y, width=w, height=h)
        entry.insert(0, self.set(row, column))
        entry.select_range(0, 'end')
        self._editor = (entry, batch_rows or (row,), column, bool(batch_rows))

        def finish_from_key(event):
            confirm = event.keysym == 'Return'
            self.finish_edit(save=confirm, apply_batch=confirm)
            self.focus_set()
            return 'break'

        entry.bind('<Return>', finish_from_key)
        entry.bind('<Escape>', finish_from_key)
        entry.bind('<Tab>', lambda _: self.finish_edit())
        entry.bind('<FocusOut>', lambda _: self.finish_edit())
        if self.paste_table and not batch_rows:
            entry.bind('<<Paste>>', self.paste_table)
        entry.focus_set()

    def finish_edit(self, save=True, *, apply_batch=False):
        if self._editor is None:
            return
        entry, rows, column, is_batch = self._editor
        self._editor = None
        changed = False
        if save and (not is_batch or apply_batch):
            value = entry.get()
            for row in rows:
                if self.exists(row) and value != self.set(row, column):
                    self.set(row, column, value)
                    changed = True
        entry.destroy()
        if changed:
            self.event_generate('<<RowEdited>>')


# ── Main App ──────────────────────────────────────────────────────────────
class App(ttk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack(fill='both', expand=True)
        self.df         = None
        self.port_data  = []
        self.sizes      = []
        self.po_number  = ''   # read from CSV ORDER NUMBER column
        # 客户色号 → 逐件英文（'|' 连接），由拆分弹窗填入，输出到「英文颜色拆分」末列
        self._en_split  = {}
        # 客户色号 → 逐件厂款号（'|' 连接），由拆分弹窗填入，输出到「厂款号拆分」末列
        self._fac_split = {}
        # 批量模式：{文件名: (合同号, df)}；非空表示 ③ 页的生成按钮走批量写出
        self._batch_pos = {}
        self._batch_dir = ''
        self._pack_colour_codes = set()
        self._erp_colour_overrides = {}
        self._erp_client = OrderErpClient()
        self._erp_credentials = CredentialStore()
        self._colour_library_snapshots = {}
        self._preview_split = False   # 预览视图：False=按色号，True=逐件拆分
        self._colour_import_changes = []
        self._colour_library_write = None
        self.out_dir_var = tk.StringVar()
        self._setup_style()
        self._build_ui()
        self._load_config()

    # ── Styles ────────────────────────────────────────────────────────────
    def _setup_style(self):
        s = ttk.Style(self.winfo_toplevel())
        s.theme_use('clam')
        s.configure('H.TLabel',          font=('Arial', 11, 'bold'))
        s.configure('Sub.TLabel',        font=('Arial', 8), foreground='gray')
        s.configure('TLabelframe.Label', font=('Arial', 9, 'bold'))
        s.configure('Accent.TButton',    font=('Arial', 10, 'bold'))
        s.configure('TNotebook.Tab',     font=('Arial', 10), padding=(14, 6))

    # ── Config persistence ────────────────────────────────────────────────
    def _load_config(self):
        try:
            with open(CONFIG_FILE, encoding='utf-8') as f:
                cfg = json.load(f)
            for attr, key in (('xlsx_var', 'xlsx_path'), ('port_var', 'port_path'),
                              ('price_var', 'price_path'),
                              ('colour_library_var', 'colour_library_path')):
                v = cfg.get(key, '')
                if v and os.path.exists(v):
                    getattr(self, attr).set(v)
            out = cfg.get('out_dir', '')
            if out:
                self.out_dir_var.set(out)
        except Exception:
            pass

    def _save_config(self):
        try:
            cfg = {}
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, encoding='utf-8') as f:
                    cfg = json.load(f)
            cfg['xlsx_path']  = self.xlsx_var.get()
            cfg['port_path']  = self.port_var.get()
            cfg['price_path'] = self.price_var.get()
            cfg['colour_library_path'] = self.colour_library_var.get().strip()
            cfg['out_dir']    = self.out_dir_var.get()
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    # ── UI build ──────────────────────────────────────────────────────────
    def _build_ui(self):
        ttk.Label(self, text='PO → 订单明细 导入工具',
                  style='H.TLabel').pack(pady=(8, 4))

        self.nb = ttk.Notebook(self)
        self.nb.pack(fill='both', expand=True, padx=10, pady=(0, 6))

        # Tabs
        self.tab_file    = ttk.Frame(self.nb, padding=12)
        self.tab_config  = ttk.Frame(self.nb, padding=12)
        self.tab_preview = ttk.Frame(self.nb, padding=12)

        self.nb.add(self.tab_file,    text='  ① 文件选择  ')
        self.nb.add(self.tab_config,  text='  ② 字段配置  ')
        self.nb.add(self.tab_preview, text='  ③ 预览 & 生成  ')

        self._build_tab_file()
        self._build_tab_config()
        self._build_tab_preview()

        # Switch to preview tab → refresh
        self.nb.bind('<<NotebookTabChanged>>', self._on_tab_change)

        # Bottom status bar (always visible)
        bar = ttk.Frame(self)
        bar.pack(fill='x', padx=10, pady=(0, 6))
        self.status = tk.StringVar(value='请在 ① 选择文件并加载数据')
        ttk.Label(bar, textvariable=self.status,
                  foreground='#005599').pack(side='left')

    # ── Tab 1: File selection ─────────────────────────────────────────────
    def _build_tab_file(self):
        p = self.tab_file
        p.columnconfigure(1, weight=1)

        rows_cfg = [
            ('PO 文件 (PDF/CSV):', 'csv_var',  self._pick_csv,  '(每次选择)'),
            ('订单明细 模板:',     'xlsx_var', self._pick_xlsx, '(已记忆路径)'),
            ('港口 文件:',         'port_var', self._pick_port, '(已记忆路径)'),
            ('下单价格表:',        'price_var', self._pick_price,
             '(已记忆路径，批量取厂款号)'),
        ]
        for r, (lbl, attr, cmd, hint) in enumerate(rows_cfg):
            pady = (0, 0) if r == 0 else (8, 0)
            ttk.Label(p, text=lbl).grid(row=r, column=0, sticky='w',
                                        padx=(0, 8), pady=pady)
            var = tk.StringVar()
            setattr(self, attr, var)
            ttk.Entry(p, textvariable=var, font=('Arial', 9)).grid(
                row=r, column=1, sticky='ew', padx=4, pady=pady)
            ttk.Button(p, text='浏览', command=cmd, width=6).grid(
                row=r, column=2, pady=pady)
            ttk.Label(p, text=hint, style='Sub.TLabel').grid(
                row=r, column=3, sticky='w', padx=(6, 0), pady=pady)

        # Export folder row (persistent, no dialog on generate)
        ttk.Label(p, text='导出文件夹:').grid(row=4, column=0, sticky='w',
                                              padx=(0, 8), pady=(8, 0))
        ttk.Entry(p, textvariable=self.out_dir_var, font=('Arial', 9)).grid(
            row=4, column=1, sticky='ew', padx=4, pady=(8, 0))
        ttk.Button(p, text='浏览', width=6,
                   command=self._pick_out_dir).grid(row=4, column=2, pady=(8, 0))
        ttk.Label(p, text='(已记忆路径)', style='Sub.TLabel').grid(
            row=4, column=3, sticky='w', padx=(6, 0), pady=(8, 0))

        ttk.Separator(p, orient='horizontal').grid(
            row=5, column=0, columnspan=4, sticky='ew', pady=16)

        btnrow = ttk.Frame(p)
        btnrow.grid(row=6, column=1, columnspan=3, sticky='w')
        ttk.Button(btnrow, text='  加载数据  ', command=self._load).pack(side='left')
        ttk.Button(btnrow, text='  批量：扫描文件夹  ',
                   command=self._batch_scan).pack(side='left', padx=(12, 0))

        ttk.Label(p,
            text='加载后将自动跳转到 ② 字段配置 页；路径字段均会自动记忆\n'
                 '批量：选一个放 PO 的文件夹（.pdf/.csv，同名时优先 CSV）→ 把全部 PO 的'
                 '客户色号合并去重后填进 ② 的颜色表（厂款号按下单价格表预填），\n'
                 '像单个 PO 那样核对颜色映射，填写工厂出厂期，再到 ③ 点「生成 订单明细」，'
                 '一次写出全部文件。',
            style='Sub.TLabel', justify='left').grid(
            row=7, column=1, columnspan=3, sticky='w', pady=(10, 0))

        self.colour_library_var = tk.StringVar()
        ttk.Label(p, text='颜色库：').grid(row=8, column=0, sticky='w',
                                           padx=(0, 8), pady=(12, 0))
        ttk.Entry(p, textvariable=self.colour_library_var, font=('Arial', 9)).grid(
            row=8, column=1, sticky='ew', padx=4, pady=(12, 0))
        ttk.Button(p, text='浏览', width=6, command=self._pick_colour_library).grid(
            row=8, column=2, pady=(12, 0))
        ttk.Label(p, text='加载时按英文名称匹配', style='Sub.TLabel').grid(
            row=8, column=3, sticky='w', padx=(6, 0), pady=(12, 0))

    # ── Tab 2: Field config ───────────────────────────────────────────────
    def _build_tab_config(self):
        p = self.tab_config

        # ---- 厂款号和工厂出厂期批量填入 ----
        extra_row = ttk.Frame(p)
        extra_row.pack(fill='x', pady=(0, 8))

        ttk.Label(extra_row, text='厂款号:').pack(side='left')
        self.fac_style_var = tk.StringVar()
        ttk.Entry(extra_row, textvariable=self.fac_style_var,
                  width=14).pack(side='left', padx=(4, 8))
        ttk.Label(extra_row, text='厂款号2:').pack(side='left')
        self.fac_style2_var = tk.StringVar()
        ttk.Entry(extra_row, textvariable=self.fac_style2_var,
                  width=14).pack(side='left', padx=(4, 8))
        ttk.Label(extra_row, text='工厂出厂期:').pack(side='left')
        self.fac_date_var = tk.StringVar()
        ttk.Entry(extra_row, textvariable=self.fac_date_var, width=12,
                  state='readonly').pack(side='left', padx=(4, 2))
        ttk.Button(extra_row, text='选日期',
                   command=self._pick_fac_date).pack(side='left', padx=(0, 8))
        ttk.Button(extra_row, text='配置',
                   command=self._fill_fac_style).pack(side='left')
        ttk.Label(extra_row,
                  text='款号按勾选填入，点击 ☑ 列头可全选或取消。日期填入整列。',
                  style='Sub.TLabel').pack(side='left', padx=(8, 0))

        # ---- 多拼色拆分：作用于选中行（也可双击色号/色名列） ----
        split_bar = ttk.Frame(p)
        split_bar.pack(fill='x', pady=(0, 8))
        ttk.Button(split_bar, text='拆分所选颜色…',
                   command=self._split_selected).pack(side='left')
        ttk.Label(split_bar,
                  text='   选中一行后点此，或双击「客户色号/色名」列，按件数把多拼色拆成逐件中英文',
                  style='Sub.TLabel').pack(side='left', padx=(8, 0))

        # ---- Colour table ----
        hdr_colour = ttk.Frame(p)
        hdr_colour.pack(fill='x', pady=(0, 3))
        ttk.Label(hdr_colour, text='颜色映射',
                  font=('Arial', 9, 'bold')).pack(side='left')
        ttk.Label(hdr_colour,
                  text='   单击 ✏ 编辑。按住 Ctrl 点选同一列的颜色名称或厂款号，按 Enter 编辑，再按 Enter 保存。',
                  style='Sub.TLabel').pack(side='left')

        colour_cols = ('cust_code', 'cust_name', 'colour_id',
                       'colour_cn', 'colour_en',
                       'fac_chk', 'fac_style', 'fac2_chk', 'fac_style2',
                       'fac_date')
        self.colour_tv = InlineEditTreeview(
            p, editable_cols=[2, 3, 4, 6, 8, 9],
            paste_table=self._paste_colour_table, batch_columns=('colour_cn', 'fac_style'),
            columns=colour_cols, show='headings',
            selectmode='extended', height=7)
        self.colour_tv.heading('cust_code', text='客户色号（自动）')
        self.colour_tv.heading('cust_name', text='客户色名（自动）')
        self.colour_tv.heading('colour_id', text='颜色编号 ✏')
        self.colour_tv.heading('colour_cn', text='颜色名称 ✏')
        self.colour_tv.heading('colour_en', text='英文颜色 ✏')
        self.colour_tv.heading('fac_chk',   text='☑',
                               command=lambda: self._toggle_all_chk(5))
        self.colour_tv.heading('fac_style', text='厂款号 ✏')
        self.colour_tv.heading('fac2_chk',  text='☑',
                               command=lambda: self._toggle_all_chk(7))
        self.colour_tv.heading('fac_style2', text='厂款号2 ✏')
        self.colour_tv.heading('fac_date',  text='工厂出厂期 ✏')
        self.colour_tv.column('cust_code', width=100, anchor='center')
        self.colour_tv.column('cust_name', width=100, anchor='center')
        self.colour_tv.column('colour_id', width=110, anchor='center')
        self.colour_tv.column('colour_cn', width=110, anchor='center')
        self.colour_tv.column('colour_en', width=130, anchor='center')
        self.colour_tv.column('fac_chk',   width=34,  anchor='center',
                              stretch=False)
        self.colour_tv.column('fac_style', width=110, anchor='center')
        self.colour_tv.column('fac2_chk',  width=34,  anchor='center',
                              stretch=False)
        self.colour_tv.column('fac_style2', width=110, anchor='center')
        self.colour_tv.column('fac_date',  width=110, anchor='center')
        self.colour_tv.pack(fill='x', pady=(0, 4))
        self.colour_tv.bind('<<RowEdited>>', lambda _: self._refresh_preview())
        self.colour_tv.bind('<Control-v>', self._paste_colour_table)
        self.colour_tv.bind('<Control-V>', self._paste_colour_table)
        self.colour_tv.bind('<Button-1>', self._on_colour_click)
        # 双击只读的「客户色号/色名」列 → 多拼色拆分弹窗
        self.colour_tv.bind('<Double-1>', self._open_colour_split)

        ttk.Label(p,
            text='Ctrl+V 粘贴：从 Excel 复制列（颜色编号、颜色名称、英文颜色、厂款号、工厂出厂期），选中起始行后粘贴',
            style='Sub.TLabel').pack(anchor='w', pady=(0, 10))

        # ---- Destination table ----
        hdr_dest = ttk.Frame(p)
        hdr_dest.pack(fill='x', pady=(0, 3))
        ttk.Label(hdr_dest, text='目的地映射',
                  font=('Arial', 9, 'bold')).pack(side='left')
        ttk.Label(hdr_dest,
                  text='   自动从港口文件查找；单击带 ✏ 的单元格可修改',
                  style='Sub.TLabel').pack(side='left')

        dest_cols = ('code', 'full', 'port_cn', 'port_code')
        self.dest_tv = InlineEditTreeview(
            p, editable_cols=[2, 3],
            columns=dest_cols, show='headings', height=8)
        self.dest_tv.heading('code',      text='代码')
        self.dest_tv.heading('full',      text='目的地（CSV 原文）')
        self.dest_tv.heading('port_cn',   text='目的港 ✏')
        self.dest_tv.heading('port_code', text='港口代码 ✏')
        self.dest_tv.column('code',      width=58,  anchor='center')
        self.dest_tv.column('full',      width=420)
        self.dest_tv.column('port_cn',   width=160, anchor='center')
        self.dest_tv.column('port_code', width=220, anchor='center')
        self.dest_tv.bind('<<RowEdited>>', lambda _: self._refresh_preview())

        import_actions = ttk.Frame(p)
        import_actions.pack(side='bottom', fill='x', pady=(8, 0))
        ttk.Button(import_actions, text='预览ERP物料颜色导入清单',
                   command=self._preview_erp_colour_list).grid(
            row=0, column=0, sticky='ew', pady=(0, 4))
        ttk.Button(import_actions, text='预览ERP成品颜色导入清单',
                   command=lambda: self._preview_erp_colour_list(finished=True)).grid(
            row=0, column=1, sticky='ew', padx=(8, 0), pady=(0, 4))
        self.erp_login_btn = ttk.Button(import_actions, text='登录罗斯系统', command=self._erp_login_dialog)
        self.erp_login_btn.grid(row=1, column=0, sticky='ew', pady=(0, 4))
        self.erp_import_btn = ttk.Button(import_actions, text='导入ERP', command=self._import_erp_colours)
        self.erp_import_btn.grid(row=2, column=0, sticky='ew', pady=(0, 4))
        self.erp_status = tk.StringVar(master=self, value='未登录')
        ttk.Label(import_actions, textvariable=self.erp_status).grid(
            row=1, column=1, sticky='w', padx=(8, 0), pady=(0, 4))
        self.erp_undo_btn = ttk.Button(import_actions, text='删除本次导入', command=self._undo_erp_import)
        self.erp_undo_btn.grid(row=2, column=1, sticky='ew', padx=(8, 0), pady=(0, 4))
        self.erp_result_btn = ttk.Button(import_actions, text='查到导入结果', command=self._show_erp_import_result)
        self.erp_result_btn.grid(row=2, column=2, sticky='ew', padx=(8, 0), pady=(0, 4))
        ttk.Button(import_actions, text='导入颜色库',
                   command=lambda: self._import_colour_library(interactive=True)).grid(
            row=3, column=0, sticky='ew')
        self.undo_colour_import_btn = ttk.Button(
            import_actions, text='删除本次导入', command=self._undo_colour_import,
            state='disabled')
        self.undo_colour_import_btn.grid(row=3, column=1, sticky='ew', padx=(8, 0))
        self.dest_tv.pack(fill='x')

        # Tk clicks do not bubble to the page frame, so bind each background surface.
        for widget in (p, *p.winfo_children()):
            if isinstance(widget, (ttk.Frame, ttk.Label)) or widget is self.dest_tv:
                widget.bind('<ButtonRelease-1>', self._clear_colour_batch_on_blank, add='+')

    def _run_erp_task(self, status, task, on_done):
        self.erp_login_btn.state(['disabled'])
        self.erp_import_btn.state(['disabled'])
        self.erp_undo_btn.state(['disabled'])
        self.erp_result_btn.state(['disabled'])
        self.erp_status.set(status)
        done = Queue(maxsize=1)

        def worker():
            try:
                done.put((task(), None))
            except Exception as ex:
                done.put((None, ex))

        def finish():
            try:
                result, error = done.get_nowait()
            except Empty:
                self.after(80, finish)
                return
            self.erp_login_btn.state(['!disabled'])
            self.erp_import_btn.state(['!disabled'])
            self.erp_undo_btn.state(['!disabled'])
            self.erp_result_btn.state(['!disabled'])
            on_done(result, error)
            self.erp_status.set('已登录' if self._erp_client.logged_in else '未登录')

        Thread(target=worker, daemon=True).start()
        self.after(80, finish)

    def _erp_login_dialog(self):
        parent = self.winfo_toplevel()
        try:
            saved = self._erp_credentials.load()
        except (CredentialError, OSError) as ex:
            messagebox.showwarning('登录罗斯系统', str(ex), parent=parent)
            saved = None
        if saved and saved['url'] != ERP_URL:
            saved = None

        win = tk.Toplevel(self)
        win.title('登录罗斯系统')
        win.transient(parent)
        win.resizable(False, False)
        frame = ttk.Frame(win, padding=14)
        frame.pack(fill='both', expand=True)
        username = tk.StringVar(master=win, value=saved['username'] if saved else '')
        password = tk.StringVar(master=win, value=saved['password'] if saved else '')
        ttk.Label(frame, text='账号').grid(row=0, column=0, sticky='w', pady=(0, 6))
        account_entry = ttk.Entry(frame, textvariable=username, width=29)
        account_entry.grid(row=0, column=1, pady=(0, 6), padx=(8, 0))
        ttk.Label(frame, text='密码').grid(row=1, column=0, sticky='w')
        password_entry = ttk.Entry(frame, textvariable=password, width=29, show='*')
        password_entry.grid(row=1, column=1, padx=(8, 0))
        ttk.Label(frame, text='登录成功后会记住账号和密码。').grid(
            row=2, column=0, columnspan=2, sticky='w', pady=(8, 0))

        def submit(_=None):
            account, secret = username.get().strip(), password.get()
            if not account or not secret:
                messagebox.showwarning('登录罗斯系统', '请输入账号和密码。', parent=win)
                return
            password.set('')
            win.destroy()

            def login():
                client = OrderErpClient(undo_directory=self._erp_credentials.path.parent)
                client.login(account, secret)
                try:
                    self._erp_credentials.save(ERP_URL, account, secret)
                except (CredentialError, OSError) as ex:
                    return client, str(ex)
                return client, ''

            def on_login(result, error):
                if error is not None:
                    messagebox.showerror('登录罗斯系统', str(error), parent=parent)
                    return
                self._erp_client, storage_error = result
                if storage_error:
                    messagebox.showwarning('登录罗斯系统', f'登录成功，但无法记住密码。\n{storage_error}', parent=parent)
                else:
                    messagebox.showinfo('登录罗斯系统', '登录成功，已记住账号和密码。', parent=parent)

            self._run_erp_task('正在登录…', login, on_login)

        actions = ttk.Frame(frame)
        actions.grid(row=3, column=0, columnspan=2, sticky='e', pady=(12, 0))
        ttk.Button(actions, text='登录', command=submit).pack(side='left', padx=(0, 8))
        ttk.Button(actions, text='取消', command=win.destroy).pack(side='left')
        win.bind('<Return>', submit)
        win.bind('<Escape>', lambda _: win.destroy())
        win.grab_set()
        (password_entry if saved else account_entry).focus_set()

    def _import_erp_colours(self):
        parent = self.winfo_toplevel()
        mapping_rows = self._erp_colour_mapping_rows()
        try:
            rows = self._current_erp_colour_rows(mapping_rows)
        except Exception as ex:
            messagebox.showwarning('颜色库', f'无法读取颜色库：\n{ex}', parent=parent)
            return
        material = self._classified_erp_colour_rows(rows, finished=False)
        finished = self._classified_erp_colour_rows(rows, finished=True)
        missing = self._erp_missing_fields(mapping_rows, material, finished)
        if missing:
            self._show_erp_missing_fields(missing)
            return
        if not rows:
            messagebox.showinfo('导入ERP', '筛选后没有可导入的色号。', parent=parent)
            return
        client = self._erp_client
        if not client.logged_in:
            messagebox.showinfo('导入ERP', '请先点击「登录罗斯系统」。', parent=parent)
            return

        def on_import(batch, error):
            if error is not None:
                messagebox.showerror('导入ERP', str(error), parent=parent)
                return
            rows = batch['rows']
            added = sum(row['status'] == 'added' for row in rows)
            existing = sum(row['status'] == 'existing' for row in rows)
            failed = len(rows) - added - existing
            notify = messagebox.showwarning if failed else messagebox.showinfo
            notify('导入ERP', f'导入成功 {added} 条，已存在 {existing} 条，导入失败 {failed} 条。\n'
                   '点击「查到导入结果」查看明细。', parent=parent)

        self._run_erp_task('正在导入…', lambda: client.import_colours(material[:2], finished[:2]), on_import)

    @staticmethod
    def _erp_missing_fields(mapping_rows, material, finished):
        missing = []
        # Blank codes are omitted from the previews, so report their source cells.
        code_fields = ((0, '客户色号'), (2, '颜色编号'))
        for number, row in enumerate(mapping_rows, 1):
            fields = [label for index, label in code_fields if not row[index].strip()]
            if fields:
                missing.append(f'颜色映射第 {number} 行缺少「{"、".join(fields)}」。')

        labels = ('色号', '颜色名称', '颜色分类', '英文名称')
        for (library, _, _, _), rows in zip(LIBRARIES, (material, finished)):
            for number, row in enumerate(rows, 1):
                fields = [label for label, value in zip(labels, row) if not value.strip()]
                if fields:
                    missing.append(f'ERP{library}导入清单第 {number} 行（{row[0]}）缺少「{"、".join(fields)}」。')
        return missing

    def _show_erp_missing_fields(self, missing):
        win = tk.Toplevel(self)
        win.title('ERP 导入清单数据不完整')
        win.transient(self.winfo_toplevel())
        frame = ttk.Frame(win, padding=12)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text='请补齐以下数据后再点击「导入ERP」。').pack(anchor='w', pady=(0, 8))
        details = ScrolledText(frame, width=68, height=min(max(len(missing), 3), 14), wrap='word')
        details.pack(fill='both', expand=True)
        details.insert('1.0', '\n'.join(missing))
        details.configure(state='disabled')
        close = ttk.Button(frame, text='关闭', command=win.destroy)
        close.pack(anchor='e', pady=(8, 0))
        win.bind('<Escape>', lambda _: win.destroy())
        win.grab_set()
        close.focus_set()

    def _show_erp_import_result(self):
        parent = self.winfo_toplevel()
        title = 'ERP 导入结果'
        try:
            batch = self._erp_client.import_result()
        except ErpError as ex:
            messagebox.showerror(title, str(ex), parent=parent)
            return
        rows = batch['rows']
        if not rows:
            text = '当前账号没有可查看的 ERP 导入结果。' if self._erp_client.logged_in else '请先点击「登录罗斯系统」。'
            messagebox.showinfo(title, text, parent=parent)
            return

        statuses = {'added': '导入成功', 'existing': '已存在', 'failed': '导入失败'}
        libraries = {module: label for label, module, _, _ in LIBRARIES}
        counts = {key: sum(row['status'] == key for row in rows) for key in statuses}
        summary = '    '.join(f'{label} {counts[key]} 条' for key, label in statuses.items())
        if batch['started_at']:
            summary = f'{batch["started_at"]}    {summary}'

        win = tk.Toplevel(self)
        win.title(title)
        win.transient(parent)
        win.grab_set()
        frame = ttk.Frame(win, padding=12)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text=summary).pack(anchor='w', pady=(0, 8))
        columns = (
            ('library', '颜色库', 100), ('code', '色号', 130), ('name', '颜色名称', 150),
            ('category', '颜色分类', 125), ('english', '英文名称', 170), ('status', '导入结果', 172),
        )
        table = ttk.Treeview(frame, columns=[key for key, _, _ in columns],
                             show='headings', selectmode='browse', height=len(rows))
        for key, label, width in columns:
            table.heading(key, text=label)
            table.column(key, width=width, minwidth=65, anchor='w')
        table.pack(fill='both', expand=True)
        details = {}
        undo_notes = {'deleted': '已撤销', 'absent': 'ERP 中已不存在'}
        for row in rows:
            status = statuses[row['status']]
            undo = undo_notes.get(row.get('undo'))
            if undo:
                status += '（' + undo + '）'
            item = table.insert('', 'end', values=(libraries[row['module']], row['code'], row['name'],
                                                  row['category'], row['english'], status))
            details[item] = row['detail'] + (f' {undo}。' if undo else '')
        detail = ttk.Label(frame, justify='left', wraplength=700)
        detail.pack(fill='x', pady=(8, 0))

        def show_detail(_=None):
            selected = table.selection()
            detail.configure(text=details[selected[0]] if selected else '')

        table.bind('<<TreeviewSelect>>', show_detail)
        table.selection_set(next(iter(details)))
        show_detail()
        ttk.Button(frame, text='关闭', command=win.destroy).pack(anchor='e', pady=(8, 0))
        win.bind('<Escape>', lambda _: win.destroy())

    def _undo_erp_import(self):
        parent = self.winfo_toplevel()
        client = self._erp_client
        title = '删除本次 ERP 导入'
        if not client.logged_in:
            messagebox.showinfo(title, '请先点击「登录罗斯系统」。', parent=parent)
            return
        try:
            records = client.undo_records()
        except ErpError as ex:
            messagebox.showerror(title, str(ex), parent=parent)
            return
        if not records:
            messagebox.showinfo(title, '当前账号没有可撤销的 ERP 导入记录。', parent=parent)
            return

        lines = []
        for label, module, _, _ in LIBRARIES:
            codes = [snapshot[0] for target, snapshot in records if target == module]
            if codes:
                lines.append(f'{label}：{"、".join(codes)}')
        if not messagebox.askyesno(title, '确认从罗斯系统删除以下色号？\n\n' + '\n'.join(lines),
                                   default=messagebox.NO, parent=parent):
            return

        def on_undo(results, error):
            if error is not None:
                messagebox.showerror(title, str(error), parent=parent)
                return
            lines = []
            incomplete = False
            for label, outcomes in results.items():
                for key, status in (('deleted', '已删除'), ('absent', '已不存在'), ('blocked', '未删除')):
                    codes = outcomes[key]
                    if codes:
                        lines.append(f'{label}{status}：{"、".join(codes)}')
                incomplete |= bool(outcomes['blocked'])
            notify = messagebox.showwarning if incomplete else messagebox.showinfo
            notify(title, '\n'.join(lines), parent=parent)

        self._run_erp_task('正在删除…', lambda: client.undo_import(records), on_undo)

    def _erp_colour_mapping_rows(self):
        mapping = self.colour_tv
        mapping.finish_edit()
        source_columns = ('cust_code', 'cust_name', 'colour_id', 'colour_cn')
        return [tuple(mapping.set(item, column) for column in source_columns)
                for item in mapping.get_children()]

    def _current_erp_colour_rows(self, mapping_rows=None):
        if mapping_rows is None:
            mapping_rows = self._erp_colour_mapping_rows()
        library_path = self.colour_library_var.get().strip()
        library_key = os.path.normcase(os.path.abspath(library_path))
        if mapping_rows and library_path and library_key not in self._colour_library_snapshots:
            self._read_colour_library(library_path)
        initial_max_co, existing_customer_codes = self._colour_library_snapshots.get(
            library_key, (None, frozenset()))
        return _erp_colour_list_rows(mapping_rows, initial_max_co, existing_customer_codes)

    def _erp_overrides(self, finished):
        return self._erp_colour_overrides.setdefault('finished' if finished else 'material', {})

    def _classified_erp_colour_rows(self, rows, *, finished):
        overrides = self._erp_overrides(finished)
        classified = []
        for row in rows:
            code, name, english = row
            category = overrides.get(row)
            if category is None:
                category = _suggest_erp_colour_category(name, english, finished=finished)
            classified.append((code, name, category, english))
        return classified

    def _preview_erp_colour_list(self, *, finished=False):
        parent = self.winfo_toplevel()
        try:
            rows = self._current_erp_colour_rows()
        except Exception as ex:
            messagebox.showwarning('颜色库', f'无法读取颜色库：\n{ex}', parent=parent)
            return
        title = 'ERP成品颜色导入清单' if finished else 'ERP物料颜色导入清单'
        if not rows:
            messagebox.showinfo(title, '筛选后没有可预览的色号。', parent=parent)
            return

        win = tk.Toplevel(self)
        win.title(title)
        win.transient(parent)
        win.minsize(580, 260)
        win.grab_set()

        ttk.Label(win, text=f'共 {len(rows)} 条。单击颜色分类单元格可展开下拉选项。').pack(
            anchor='w', padx=12, pady=(10, 6))
        ttk.Button(win, text='关闭', command=win.destroy).pack(side='bottom', pady=8)
        frame = ttk.Frame(win)
        frame.pack(fill='both', expand=True, padx=12)
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)
        table = ttk.Treeview(frame, columns=('code', 'name', 'category', 'english'),
                             show='headings', selectmode='browse', height=min(len(rows), 16))
        for column, label, width in (
            ('code', '色号', 125), ('name', '颜色名称', 175),
            ('category', '颜色分类', 115), ('english', '英文名称', 260),
        ):
            table.heading(column, text=label)
            table.column(column, width=width, anchor='w', stretch=column in ('name', 'english'))
        table.grid(row=0, column=0, sticky='nsew')
        scrollbar = ttk.Scrollbar(frame, orient='vertical', command=table.yview)
        scrollbar.grid(row=0, column=1, sticky='ns')

        category_options = ERP_FINISHED_COLOUR_CATEGORIES if finished else ERP_COLOUR_CATEGORIES
        overrides = self._erp_overrides(finished)
        row_keys = {}
        categories = {}
        for code, name, category, english in self._classified_erp_colour_rows(rows, finished=finished):
            item = table.insert('', 'end', values=(code, name, f'{category or "请选择"}  ▾', english))
            row_keys[item] = (code, name, english)
            categories[item] = category

        editor = None

        def close_editor():
            nonlocal editor
            if editor is not None:
                editor.destroy()
                editor = None

        def on_scroll(first, last):
            close_editor()
            scrollbar.set(first, last)

        def edit_category(event):
            nonlocal editor
            close_editor()
            if table.identify_region(event.x, event.y) != 'cell':
                return
            if table.identify_column(event.x) != '#3':
                return
            item = table.identify_row(event.y)
            x, y, width, height = table.bbox(item, 'category')
            combo = ttk.Combobox(table, values=category_options, state='readonly')
            editor = combo
            combo.set(categories[item])
            combo.place(x=x, y=y, width=width, height=height)

            def save_category(_):
                category = combo.get()
                categories[item] = category
                overrides[row_keys[item]] = category
                table.set(item, 'category', f'{category}  ▾')
                close_editor()
                table.focus_set()

            combo.bind('<<ComboboxSelected>>', save_category)

            def open_options():
                if editor is combo:
                    combo.event_generate('<ButtonPress-1>', x=width - 6, y=height // 2)

            combo.after_idle(open_options)

        table.configure(yscrollcommand=on_scroll)
        table.bind('<ButtonRelease-1>', edit_category)
        table.bind('<Configure>', lambda _: close_editor())

    def _clear_colour_batch_on_blank(self, event):
        if event.widget is self.dest_tv:
            if self.dest_tv.identify_region(event.x, event.y) != 'nothing':
                return
        self.colour_tv._cancel_batch_selection()

    # ── Tab 3: Preview & generate ─────────────────────────────────────────
    def _build_tab_preview(self):
        p = self.tab_preview

        self.preview_inner = ttk.Frame(p)
        self.preview_inner.pack(fill='both', expand=True)
        self.prev_tv = None
        self._rebuild_preview_tv()

        self.fname_note = tk.StringVar(value='')
        ttk.Label(p, textvariable=self.fname_note,
                  foreground='#005599').pack(anchor='w', pady=(6, 0))

        btn_bar = ttk.Frame(p)
        btn_bar.pack(fill='x', pady=(8, 0))
        ttk.Button(btn_bar, text='刷新预览',
                   command=self._refresh_preview).pack(side='left', padx=(0, 12))
        self.view_btn = ttk.Button(btn_bar, text=self._view_btn_text(),
                                   command=self._toggle_preview_view)
        self.view_btn.pack(side='left', padx=(0, 12))
        ttk.Button(btn_bar, text='  生成 订单明细  ',
                   style='Accent.TButton',
                   command=self._generate).pack(side='right')
        self.gen_note = tk.StringVar(value='')
        ttk.Label(btn_bar, textvariable=self.gen_note,
                  foreground='gray', style='Sub.TLabel').pack(side='left')

    def _on_tab_change(self, event):
        idx = self.nb.index(self.nb.select())
        if idx == 2:   # preview tab
            self.colour_tv.finish_edit()
            self.dest_tv.finish_edit()
            self._refresh_preview()

    # ── Preview view toggle（按色号 / 逐件拆分）──────────────────────────
    def _view_btn_text(self):
        return '切换：按色号视图' if self._preview_split else '切换：逐件拆分视图'

    def _toggle_preview_view(self):
        self._preview_split = not self._preview_split
        if hasattr(self, 'view_btn'):
            self.view_btn.config(text=self._view_btn_text())
        self._rebuild_preview_tv()
        self._refresh_preview()

    # 预览列名 → 行字典键（差异项；其余同名）
    _PREVIEW_KEY = {
        '实际出厂': '实际出厂（年-月-日）',
        '客户船期': '客户船期（年-月-日）',
    }

    # ── Dynamic preview treeview ──────────────────────────────────────────
    def _rebuild_preview_tv(self):
        for w in self.preview_inner.winfo_children():
            w.destroy()
        fixed = ['流水号', '厂款号', '厂款号2', '颜色编号', '颜色名称', '英文颜色',
                 '目的港', '港口代码', '客户色号', '客户色名',
                 '实际出厂', '客户船期', '数量合计']
        if not self._preview_split:
            # 按色号视图：在「英文颜色」后插入「英文颜色拆分」，「厂款号」后插入「厂款号拆分」
            i = fixed.index('英文颜色') + 1
            fixed = fixed[:i] + ['英文颜色拆分'] + fixed[i:]
            j = fixed.index('厂款号') + 1
            fixed = fixed[:j] + ['厂款号拆分'] + fixed[j:]
        prev_cols = fixed + self.sizes
        self._prev_cols = prev_cols
        self.prev_tv = ttk.Treeview(
            self.preview_inner, columns=prev_cols, show='headings', height=14)
        for col in prev_cols:
            w = 44 if col in self.sizes else (
                120 if col in ('英文颜色拆分', '厂款号拆分') else 82)
            self.prev_tv.heading(col, text=col)
            self.prev_tv.column(col, width=w, anchor='center', minwidth=36)
        ybar = ttk.Scrollbar(self.preview_inner, orient='vertical',
                             command=self.prev_tv.yview)
        xbar = ttk.Scrollbar(self.preview_inner, orient='horizontal',
                             command=self.prev_tv.xview)
        self.prev_tv.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        ybar.pack(side='right', fill='y')
        xbar.pack(side='bottom', fill='x')
        self.prev_tv.pack(fill='both', expand=True)

    # ── File pickers ──────────────────────────────────────────────────────
    def _pick_csv(self):
        p = filedialog.askopenfilename(
            title='选择 PO 文件（PDF / CSV）',
            filetypes=[('PO 文件', '*.pdf *.csv'), ('PDF 文件', '*.pdf'),
                       ('CSV 文件', '*.csv'), ('所有文件', '*.*')])
        if p:
            self.csv_var.set(p)
            candidate = os.path.join(os.path.dirname(p), '港口.xlsx')
            if os.path.exists(candidate) and not self.port_var.get():
                self.port_var.set(candidate)

    def _pick_price(self):
        p = filedialog.askopenfilename(
            title='选择 CO-下单价格表',
            filetypes=[('Excel 文件', '*.xlsx *.xls'), ('所有文件', '*.*')])
        if p:
            self.price_var.set(p)
            self._save_config()

    def _pick_xlsx(self):
        p = filedialog.askopenfilename(
            title='选择订单明细模板',
            filetypes=[('Excel 文件', '*.xlsx'), ('所有文件', '*.*')])
        if p:
            self.xlsx_var.set(p)
            self._save_config()

    def _pick_out_dir(self):
        d = filedialog.askdirectory(title='选择导出文件夹')
        if d:
            self.out_dir_var.set(d)
            self._save_config()

    def _pick_port(self):
        p = filedialog.askopenfilename(
            title='选择港口文件',
            filetypes=[('Excel 文件', '*.xlsx'), ('所有文件', '*.*')])
        if p:
            self.port_var.set(p)
            self._save_config()

    def _pick_colour_library(self):
        path = filedialog.askopenfilename(
            title='选择颜色库文件',
            filetypes=[('Excel 文件', '*.xlsx *.xls *.xlsm'), ('所有文件', '*.*')])
        if path:
            self.colour_library_var.set(path)
            self._save_config()

    # ── Port lookup ───────────────────────────────────────────────────────
    def _load_port_data(self):
        path = self.port_var.get().strip()
        if not path or not os.path.exists(path):
            self.port_data = []
            return
        try:
            self.port_data = pd.read_excel(path, header=0).to_dict('records')
        except Exception as ex:
            messagebox.showwarning('港口文件', f'无法读取港口文件：{ex}')
            self.port_data = []

    def _lookup_port(self, dest_name):
        dest_code = dest_name.split(' - ')[0].strip()
        for entry in self.port_data:
            co = str(entry.get('对应CO合同港口') or '')
            if co and co != 'nan' and co.split(' - ')[0].strip() == dest_code:
                return (str(entry.get('名称') or ''),
                        str(entry.get('英文名称') or ''))
        return '', ''

    # ── Load ─────────────────────────────────────────────────────────────
    def _load(self):
        path = self.csv_var.get().strip()
        if not path or not os.path.exists(path):
            messagebox.showerror('错误', '请先选择有效的 PO 文件（PDF / CSV）'); return
        try:
            self.df = load_po(path)                 # .pdf 走解析，其余按 CSV
        except Exception as ex:
            messagebox.showerror('读取失败', str(ex)); return
        self._batch_pos = {}                        # 单个 PO 加载 → 退出批量模式

        needed = {'STYLE COLOUR CODE', 'COLOUR DESCRIPTION', 'DESTINATION NAME',
                  'SIZE', 'QUANTITY', 'LATEST SHIP-BY DATE',
                  'FINAL HANDOVER DATE', 'FREIGHT MODE'}
        missing = needed - set(self.df.columns)
        if missing:
            messagebox.showwarning('列缺失',
                f'CSV 中未找到以下列，相关字段将为空：\n{", ".join(sorted(missing))}')

        # Dynamic sizes — sorted small → large
        if 'SIZE' in self.df.columns:
            self.sizes = sorted(
                self.df['SIZE'].astype(str).str.strip().unique().tolist(),
                key=_size_sort_key)
        else:
            self.sizes = []

        # Capture PO number for auto filename
        self.po_number = (str(self.df['ORDER NUMBER'].iloc[0]).strip()
                          if 'ORDER NUMBER' in self.df.columns else '')

        self._pack_colour_codes = (self._csv_pack_colour_codes(self.df)
                                   if path.lower().endswith('.csv') else set())
        self._load_port_data()
        self._populate_colours()
        self._populate_dests()
        self._rebuild_preview_tv()
        self._save_config()

        nc = (self.df['STYLE COLOUR CODE'].nunique()
              if 'STYLE COLOUR CODE' in self.df.columns else 0)
        nd = (self.df['DESTINATION NAME'].nunique()
              if 'DESTINATION NAME'  in self.df.columns else 0)
        self.status.set(
            f'已加载 {len(self.df)} 行 · {nc} 种颜色 · {nd} 个目的地 · '
            f'尺码：{", ".join(self.sizes)}')

        # Check for split-file scenario and warn user upfront
        note = self._split_note()
        self.gen_note.set(note)

        # Jump to config tab
        self.nb.select(1)

    def _file_groups(self):
        """
        Return a list of lists-of-style-codes, one list per output file.
        File 1 = first style code for every colour description.
        File 2 = second style code for those descriptions that have 2+ codes.
        File 3 = third, etc.
        Single-element list means one output file (no split needed).
        """
        if self.df is None or \
           'STYLE COLOUR CODE'  not in self.df.columns or \
           'COLOUR DESCRIPTION' not in self.df.columns:
            return [[]]   # nothing

        # Build ordered {desc: [code, ...]} preserving first-appearance order
        desc_codes: dict = {}
        for code, desc in zip(self.df['STYLE COLOUR CODE'],
                               self.df['COLOUR DESCRIPTION']):
            code, desc = str(code), str(desc)
            if desc not in desc_codes:
                desc_codes[desc] = []
            if code not in desc_codes[desc]:
                desc_codes[desc].append(code)

        max_files = max((len(v) for v in desc_codes.values()), default=1)

        groups = []
        for i in range(max_files):
            grp = [codes[i] for codes in desc_codes.values() if i < len(codes)]
            if grp:
                groups.append(grp)
        return groups

    def _split_note(self):
        """Return a note string if output will be split into multiple files."""
        groups = self._file_groups()
        n = len(groups)
        if n > 1:
            return (f'⚠ 检测到颜色名称重复，将生成 {n} 个文件（'
                    f'第2份起包含重复颜色名称的第2个色号）')
        return ''

    @staticmethod
    def _sanitize(name):
        """Strip filesystem-illegal chars from a filename stem."""
        for ch in '\\/:*?"<>|':
            name = name.replace(ch, '')
        return name.strip()

    def _file_stem(self, rows):
        """Filename (no extension) for one output file, from its rows' 厂款号/厂款号2."""
        codes = []
        for r in rows:
            for key in ('厂款号', '厂款号2'):
                v = r.get(key)
                v = str(v).strip() if v else ''
                if v and v not in codes:
                    codes.append(v)
        fac = '+'.join(codes)
        stem = '_'.join(filter(None, [self.po_number, fac])) or '订单明细'
        return self._sanitize(stem)

    def _planned_outputs(self):
        """Return [(filename_without_ext, rows), ...] for every file to be written."""
        out, used = [], set()
        for style_codes in self._file_groups():
            rows = self._build_rows(style_codes=set(style_codes))
            if not rows:
                continue
            for j, r in enumerate(rows):
                r['流水号'] = j + 1          # restart 流水号 per file
            stem = self._file_stem(rows)
            name, k = stem, 2
            while name in used:              # guard against duplicate stems
                name, k = f'{stem}_{k}', k + 1
            used.add(name)
            out.append((name, rows))
        return out

    # Checkbox columns in colour table: 5 = 厂款号☑, 7 = 厂款号2☑
    _CHK_COLS = (5, 7)

    def _on_colour_click(self, event):
        """Toggle ☑/☐ when a checkbox cell is clicked."""
        if self.colour_tv.identify('region', event.x, event.y) != 'cell':
            return
        col_idx = int(self.colour_tv.identify_column(event.x).replace('#', '')) - 1
        if col_idx not in self._CHK_COLS:
            return
        row_id = self.colour_tv.identify_row(event.y)
        if not row_id:
            return
        vals = list(self.colour_tv.item(row_id, 'values'))
        vals[col_idx] = '☐' if vals[col_idx] == '☑' else '☑'
        self.colour_tv.item(row_id, values=vals)
        return 'break'

    def _toggle_all_chk(self, col_idx):
        """Heading click: check all rows in this column, or uncheck if all checked."""
        items = self.colour_tv.get_children()
        if not items:
            return
        all_checked = all(
            self.colour_tv.item(i, 'values')[col_idx] == '☑' for i in items)
        new = '☐' if all_checked else '☑'
        for i in items:
            vals = list(self.colour_tv.item(i, 'values'))
            vals[col_idx] = new
            self.colour_tv.item(i, values=vals)

    def _pick_fac_date(self):
        today = date.today()
        selected = self.fac_date_var.get()
        initial = date.fromisoformat(selected) if selected else today
        year, month = initial.year, initial.month
        calendar = Calendar()

        win = tk.Toplevel(self)
        win.title('选择工厂出厂期')
        win.transient(self.winfo_toplevel())
        win.resizable(False, False)
        win.bind('<Escape>', lambda _: win.destroy())

        panel = ttk.Frame(win, padding=8)
        panel.pack()
        for col in range(7):
            panel.columnconfigure(col, weight=1, uniform='day')
        month_label = ttk.Label(panel, anchor='center')
        month_label.grid(row=0, column=1, columnspan=5, sticky='ew')
        for col, weekday in enumerate('一二三四五六日'):
            ttk.Label(panel, text=weekday, anchor='center').grid(
                row=1, column=col, sticky='ew', pady=(4, 2))

        def choose(value):
            self.fac_date_var.set(value)
            win.destroy()

        def show_month():
            month_label.configure(text=f'{year}年{month}月')
            weeks = calendar.monthdayscalendar(year, month)
            for index, button in enumerate(day_buttons):
                week, col = divmod(index, 7)
                day = weeks[week][col] if week < len(weeks) else 0
                if day:
                    value = date(year, month, day).isoformat()
                    button.configure(text=str(day), state='normal',
                                     command=lambda value=value: choose(value))
                else:
                    button.configure(text='', state='disabled')

        def shift_month(step):
            nonlocal year, month
            next_year, month_index = divmod(year * 12 + month - 1 + step, 12)
            if 1 <= next_year <= 9999:
                year, month = next_year, month_index + 1
                show_month()

        ttk.Button(panel, text='‹', width=3, command=lambda: shift_month(-1)).grid(
            row=0, column=0, sticky='ew')
        ttk.Button(panel, text='›', width=3, command=lambda: shift_month(1)).grid(
            row=0, column=6, sticky='ew')
        day_buttons = []
        for index in range(42):
            row, col = divmod(index, 7)
            button = ttk.Button(panel, width=3)
            button.grid(row=2 + row, column=col, sticky='ew')
            day_buttons.append(button)

        ttk.Button(panel, text='今天', command=lambda: choose(today.isoformat())).grid(
            row=8, column=0, columnspan=3, sticky='w', pady=(6, 0))
        ttk.Button(panel, text='清空', command=lambda: choose('')).grid(
            row=8, column=4, columnspan=3, sticky='e', pady=(6, 0))
        show_month()
        win.wait_visibility()
        win.grab_set()
        win.focus_set()

    def _fill_fac_style(self):
        """厂款号按勾选填入，已选择的工厂出厂期填入所有行。"""
        self.colour_tv.finish_edit()
        v1 = self.fac_style_var.get().strip()
        v2 = self.fac_style2_var.get().strip()
        fac_date = self.fac_date_var.get()
        for item in self.colour_tv.get_children():
            vals = list(self.colour_tv.item(item, 'values'))
            if vals[5] == '☑':
                vals[6] = v1
            if vals[7] == '☑':
                vals[8] = v2
            if fac_date:
                vals[9] = fac_date
            self.colour_tv.item(item, values=vals)
        self._refresh_preview()

    def _split_selected(self):
        """对选中的颜色行打开逐件拆分弹窗（多选时取第一行；也可双击色号/色名列）。"""
        sel = self.colour_tv.selection()
        if not sel:
            messagebox.showinfo('拆分颜色', '请先在下方颜色表中选中一行，再点「拆分所选颜色…」。')
            return
        self._colour_split_dialog(sel[0])

    @staticmethod
    def _csv_pack_colour_codes(df):
        if not {'STYLE COLOUR CODE', 'STYLE DESCRIPTION'}.issubset(df.columns):
            return set()
        descriptions = df['STYLE DESCRIPTION'].astype('string').str.rstrip()
        is_pack = descriptions.str.contains(r' (?:SET|[0-9]*[1-9]PK)$', case=False, na=False)
        return set(df.loc[is_pack, 'STYLE COLOUR CODE'].dropna().astype(str).str.strip())

    def _populate_colours(self):
        self._en_split = {}                             # 换文件时清空逐件英文
        self._fac_split = {}                             # 换文件时清空逐件厂款号
        self._erp_colour_overrides.clear()
        self._colour_library_snapshots.clear()
        self.colour_tv.finish_edit(save=False)
        self.colour_tv._cancel_batch_selection()
        self._colour_import_changes.clear()
        self._colour_library_write = None
        self.undo_colour_import_btn.state(['disabled'])
        for item in self.colour_tv.get_children():
            self.colour_tv.delete(item)
        if self.df is None or 'STYLE COLOUR CODE' not in self.df.columns:
            return
        seen = []
        for code, en in zip(self.df['STYLE COLOUR CODE'],
                            self.df.get('COLOUR DESCRIPTION',
                                        pd.Series([''] * len(self.df)))):
            code = '' if pd.isna(code) else str(code)
            if code not in [s for s, _ in seen]:
                seen.append((code, '' if pd.isna(en) else str(en)))
        fac  = self.fac_style_var.get().strip()
        fac2 = self.fac_style2_var.get().strip()
        for cust_code, cust_name in seen:
            # CSV 中标记的套装先填客户色号，跳过颜色库英文名称匹配。
            colour_id = cust_code if cust_code.strip() in self._pack_colour_codes else ''
            self.colour_tv.insert('', 'end',
                values=(cust_code, cust_name, colour_id, '', cust_name,
                        '☑', fac, '☑', fac2, ''))
        self._import_colour_library()

    def _import_colour_library(self, *, interactive=False):
        table = self.colour_tv
        items = table.get_children()
        if not items:
            if interactive:
                messagebox.showinfo('颜色库', '请先在「① 文件选择」加载订单。')
            return
        if interactive and not self.colour_library_var.get().strip():
            self._pick_colour_library()
        if not self.colour_library_var.get().strip():
            return

        table.finish_edit()
        before = {(item, column): table.set(item, column)
                  for item in items for column in ('colour_id', 'colour_cn')}
        matched = self._apply_colour_library()
        library_write = None
        if interactive and matched:
            sheet, customer_codes, new_groups, pack_groups = matched
            if customer_codes or new_groups or pack_groups:
                path = os.path.abspath(self.colour_library_var.get().strip())
                try:
                    library_write = self._write_colour_library(
                        path, sheet, customer_codes, new_groups, pack_groups)
                except Exception as ex:
                    for (item, column), original in before.items():
                        table.set(item, column, original)
                    messagebox.showwarning('颜色库', f'无法写入颜色库：\n{ex}')
                    return
        changes = []
        for (item, column), original in before.items():
            imported = table.set(item, column)
            if imported != original:
                changes.append((item, column, original, imported))
        if changes or library_write:
            if library_write is None and self._colour_library_write is not None:
                # No new workbook write. Keep the previous write available to undo.
                combined = {(item, column): (original, imported)
                            for item, column, original, imported in self._colour_import_changes}
                for item, column, original, imported in changes:
                    first, previous = combined.get((item, column), (original, original))
                    combined[item, column] = (first if previous == original else original, imported)
                changes = [(item, column, original, imported)
                           for (item, column), (original, imported) in combined.items()]
            self._colour_import_changes = changes
            if library_write is not None:
                self._colour_library_write = library_write
            self.undo_colour_import_btn.state(['!disabled'])
            if interactive:
                self._refresh_preview()

    def _undo_colour_import(self):
        table = self.colour_tv
        table.finish_edit()
        if self._colour_library_write:
            try:
                self._undo_colour_library_write(*self._colour_library_write)
            except Exception as ex:
                messagebox.showwarning('颜色库', f'无法撤销本次颜色库导入：\n{ex}')
                return
            self._colour_library_write = None
        for item, column, original, imported in self._colour_import_changes:
            if table.exists(item) and table.set(item, column) == imported:
                table.set(item, column, original)
        self._colour_import_changes.clear()
        self.undo_colour_import_btn.state(['disabled'])
        self._refresh_preview()

    @staticmethod
    def _colour_name_key(name):
        return ' '.join(name.split()).casefold()

    @staticmethod
    def _is_colour_library_pack_row(jh_code, row_codes, pack_codes, pack_ids):
        return bool(row_codes & pack_codes) and (
            jh_code in pack_codes or (jh_code in pack_ids and row_codes <= pack_codes))

    def _read_colour_library(self, path):
        columns = ['英文名称', 'JH色号', 'JH颜色名称']
        with pd.ExcelFile(path) as book:
            for sheet in book.sheet_names:
                headers = book.parse(sheet, nrows=0).columns
                header_map = {str(header).strip(): header for header in headers}
                if all(name in header_map for name in columns):
                    source_columns = [header_map[name] for name in columns]
                    code_columns = [header for header in headers
                                    if re.fullmatch(r'(?:客人色号|客户色号)(?:\.[0-9]+)?',
                                                    str(header).strip())]
                    read_columns = source_columns + code_columns
                    data = book.parse(sheet, usecols=read_columns, dtype=str, keep_default_na=False)
                    break
            else:
                raise ValueError('找不到包含「英文名称」「JH色号」「JH颜色名称」的工作表。')

        jh_column = header_map['JH色号']
        jh_codes = data[jh_column].str.strip()
        digits = (match.group(1) for code in jh_codes if (match := CO_CODE_PATTERN.fullmatch(code)))
        max_co_number, co_width = max(((int(value), len(value)) for value in digits), default=(None, 0))

        # 保存本批首次读取时的筛选依据，后续导入写回的色号属于本批新增。
        library_key = os.path.normcase(os.path.abspath(path))
        if library_key not in self._colour_library_snapshots:
            existing_customer_codes = frozenset(
                code.casefold() for column in (jh_column, *code_columns)
                for code in data[column].str.strip() if code.startswith('6'))
            self._colour_library_snapshots[library_key] = (max_co_number, existing_customer_codes)

        return sheet, data[read_columns].itertuples(index=False, name=None), max_co_number, co_width

    def _apply_colour_library(self):
        path = self.colour_library_var.get().strip()
        items = self.colour_tv.get_children()
        if not path or not items:
            return

        packs_by_name = {}
        for item in items:
            customer_code = self.colour_tv.set(item, 'cust_code').strip()
            if customer_code in self._pack_colour_codes:
                key = self._colour_name_key(self.colour_tv.set(item, 'cust_name'))
                codes, ids = packs_by_name.setdefault(key, (set(), set()))
                codes.add(customer_code.casefold())
                colour_id = self.colour_tv.set(item, 'colour_id').strip()
                ids.add((colour_id or customer_code).casefold())

        try:
            sheet, records, max_co_number, co_width = self._read_colour_library(path)
        except Exception as ex:
            messagebox.showwarning('颜色库', f'无法读取颜色库：\n{ex}')
            return

        by_name = {}
        used_codes = set()
        for row, (english, code, name, *row_codes) in enumerate(records, start=2):
            code = code.strip()
            if code:
                used_codes.add(code.casefold())
            key = self._colour_name_key(english)
            if key:
                pack_codes, pack_ids = packs_by_name.get(key, (set(), set()))
                row_codes = {value.strip().casefold() for value in row_codes if value.strip()}
                if not self._is_colour_library_pack_row(code.casefold(), row_codes, pack_codes, pack_ids):
                    by_name.setdefault(key, (row, code, name.strip()))

        missing_ids = []
        assigned_by_name = {}
        customer_codes = defaultdict(dict)
        new_groups = {}
        pack_groups = {}

        def add_new_colour(groups, key, cust_name, item):
            if key:
                _, group_items = groups.setdefault(key, (cust_name.strip(), []))
                group_items.append(item)

        for item in items:
            cust_name = self.colour_tv.set(item, 'cust_name')
            customer_code = self.colour_tv.set(item, 'cust_code').strip()
            key = self._colour_name_key(cust_name)
            if customer_code in self._pack_colour_codes:
                code = self.colour_tv.set(item, 'colour_id').strip()
                if not code:
                    code = customer_code
                    self.colour_tv.set(item, 'colour_id', code)
                used_codes.add(code.casefold())
                add_new_colour(pack_groups, key, cust_name, item)
                continue

            if key in by_name:
                row, code, name = by_name[key]
                if customer_code:
                    customer_codes[row][customer_code] = key
                if code:
                    self.colour_tv.set(item, 'colour_id', code)
                self.colour_tv.set(item, 'colour_cn', name)
            else:
                add_new_colour(new_groups, key, cust_name, item)
            code = self.colour_tv.set(item, 'colour_id').strip()
            if code:
                used_codes.add(code.casefold())
                if key and key not in by_name and CO_CODE_PATTERN.fullmatch(code):
                    assigned_by_name.setdefault(key, code)
            else:
                missing_ids.append((item, key))

        if not missing_ids:
            return sheet, customer_codes, new_groups, pack_groups
        if max_co_number is None:
            messagebox.showwarning('颜色库', '「JH色号」列没有 CO 开头且后面全为数字的色号，无法续编。')
            return sheet, customer_codes, new_groups, pack_groups

        number = max_co_number
        for item, name_key in missing_ids:
            code = assigned_by_name.get(name_key)
            if code is None:
                while True:
                    number += 1
                    code = 'CO' + str(number).zfill(co_width)
                    if code.casefold() not in used_codes:
                        break
                used_codes.add(code.casefold())
                if name_key:
                    assigned_by_name[name_key] = code
            self.colour_tv.set(item, 'colour_id', code)
        return sheet, customer_codes, new_groups, pack_groups

    @staticmethod
    def _save_colour_library(book, path, original_stat):
        # Save beside the source so a failed write cannot truncate the workbook.
        with NamedTemporaryFile(dir=os.path.dirname(path), suffix=os.path.splitext(path)[1],
                                delete=False) as temp:
            temp_path = temp.name
        try:
            book.save(temp_path)
            current = os.stat(path)
            if (current.st_ino, current.st_size, current.st_mtime_ns) != (
                    original_stat.st_ino, original_stat.st_size, original_stat.st_mtime_ns):
                raise ValueError('颜色库已被修改，请重新导入。')
            os.replace(temp_path, path)
        finally:
            if os.path.exists(temp_path):
                os.unlink(temp_path)

    @staticmethod
    def _colour_library_cell_state(cell):
        return (cell.value, cell.data_type if cell.value is not None else 'n',
                copy(cell._style), copy(cell.comment), copy(cell.hyperlink))

    @staticmethod
    def _restore_colour_library_cell(cell, state):
        value, data_type, style, comment, hyperlink = state
        cell.value = value
        cell.data_type = data_type
        cell._style = copy(style)
        cell.comment = copy(comment)
        cell.hyperlink = copy(hyperlink)

    def _colour_library_row_state(self, cells):
        return [self._colour_library_cell_state(cell) for cell in cells]

    def _write_colour_library(self, path, sheet, customer_codes, new_groups, pack_groups):
        extension = os.path.splitext(path)[1].lower()
        if extension not in ('.xlsx', '.xlsm'):
            raise ValueError('写入颜色库需要 .xlsx 或 .xlsm 格式。')
        original_stat = os.stat(path)
        book = load_workbook(path, keep_vba=extension == '.xlsm')
        try:
            ws = book[sheet]
            headers = {str(cell.value).strip(): cell.column for cell in ws[1]}
            english_column = headers['英文名称']
            if (new_groups or pack_groups) and '中文颜色名称' not in headers:
                raise ValueError('颜色库缺少「中文颜色名称」列，无法添加新颜色。')
            customer_columns = [cell.column for cell in ws[1]
                                if cell.column > english_column
                                and str(cell.value).strip() in ('客人色号', '客户色号')]
            originals = {}

            def remember(cell):
                position = (cell.row, cell.column)
                if position not in originals:
                    originals[position] = self._colour_library_cell_state(cell)

            def write(cell, value):
                if cell.value == value:
                    return
                remember(cell)
                cell.value = value
                if value is not None:
                    cell.data_type = 's'
                cell.fill = COLOUR_IMPORT_FILL

            def append_codes(row, codes):
                cells = [ws.cell(row, column) for column in customer_columns]
                existing = {str(cell.value).strip().casefold() for cell in cells
                            if cell.value not in (None, '')}
                last_used = max((i for i, cell in enumerate(cells)
                                 if cell.value not in (None, '')), default=-1)
                blanks = iter(cells[last_used + 1:])
                for code in codes:
                    key = code.casefold()
                    if not key or key in existing:
                        continue
                    cell = next(blanks, None)
                    if cell is None:
                        column = (customer_columns[-1] if customer_columns else english_column) + 1
                        while any(ws.cell(r, column).value not in (None, '')
                                  for r in range(1, ws.max_row + 1)):
                            column += 1
                        write(ws.cell(1, column), '客人色号')
                        customer_columns.append(column)
                        cell = ws.cell(row, column)
                    write(cell, code)
                    existing.add(key)

            for row, codes in customer_codes.items():
                english = self._colour_name_key(str(ws.cell(row, english_column).value or ''))
                if any(english != name_key for name_key in codes.values()):
                    raise ValueError('颜色库行已变动，请重新导入。')
                append_codes(row, codes)

            table = self.colour_tv
            jh_column = headers['JH色号']
            # Ignore rows that contain only formatting when adding new records.
            last_data_row = ws.max_row
            while last_data_row > 1 and not any(
                    cell.value not in (None, '') for cell in ws[last_data_row]):
                last_data_row -= 1

            rows_by_name = defaultdict(list)
            for row in range(2, last_data_row + 1):
                name_key = self._colour_name_key(str(ws.cell(row, english_column).value or ''))
                if name_key in pack_groups:
                    rows_by_name[name_key].append(row)

            def add_colour(row, cust_name, items, jh_code):
                if jh_code:
                    write(ws.cell(row, jh_column), jh_code)
                write(ws.cell(row, english_column), cust_name)
                names = (table.set(item, 'colour_cn').strip() for item in items)
                colour_name = next((name for name in names if name), '')
                if colour_name:
                    for label in ('JH颜色名称', '中文颜色名称'):
                        write(ws.cell(row, headers[label]), colour_name)

            ordinary_codes_by_name = defaultdict(set)
            for codes in customer_codes.values():
                for code, name_key in codes.items():
                    ordinary_codes_by_name[name_key].add(code.casefold())
            for name_key, (_, items) in new_groups.items():
                ordinary_codes_by_name[name_key].update(
                    table.set(item, 'cust_code').strip().casefold() for item in items)

            pack_rows = set()
            for is_pack, groups in ((False, new_groups), (True, pack_groups)):
                for name_key, (cust_name, items) in groups.items():
                    codes = [table.set(item, 'cust_code').strip() for item in items]
                    numbers = [table.set(item, 'colour_id').strip() for item in items]
                    jh_code = next((number for number in numbers if number), '')
                    target = None
                    if is_pack:
                        pack_codes = {code.casefold() for code in codes}
                        pack_ids = {number.casefold() for number in numbers if number}
                        for row in rows_by_name[name_key]:
                            cells = (ws.cell(row, col) for col in customer_columns)
                            existing = {cell: str(cell.value).strip().casefold()
                                        for cell in cells if cell.value not in (None, '')}
                            existing_jh = str(ws.cell(row, jh_column).value or '').strip().casefold()
                            if self._is_colour_library_pack_row(
                                    existing_jh, set(existing.values()), pack_codes, pack_ids):
                                target = row
                                pack_rows.add(row)
                                for cell, code in existing.items():
                                    if code in ordinary_codes_by_name[name_key]:
                                        write(cell, None)
                            else:
                                for cell, code in existing.items():
                                    if code in pack_codes:
                                        write(cell, None)

                    if target is None:
                        last_data_row += 1
                        target = last_data_row
                        add_colour(target, cust_name, items, jh_code)
                    append_codes(target, codes)
                    if is_pack:
                        pack_rows.add(target)

            if pack_rows:
                rows = [row for row in range(min(pack_rows), last_data_row + 1)
                        if any(cell.value not in (None, '') for cell in ws[row])]
                ordered = [row for row in rows if row not in pack_rows] + sorted(pack_rows)
                moves = [(dest, self._colour_library_row_state(ws[src]))
                         for dest, src in zip(rows, ordered) if dest != src]
                for row, states in moves:
                    for cell, state in zip(ws[row], states):
                        if self._colour_library_cell_state(cell) == state:
                            continue
                        remember(cell)
                        value_changed = (cell.value, cell.data_type) != state[:2]
                        self._restore_colour_library_cell(cell, state)
                        if value_changed:
                            cell.fill = COLOUR_IMPORT_FILL

            for row in ws.iter_rows(max_row=last_data_row):
                for cell in row:
                    if cell.value is not None and (
                            cell.font.name != OUTPUT_FONT_NAME or cell.font.scheme is not None):
                        remember(cell)
                        _set_output_font(cell)

            written = []
            for (row, column), before in originals.items():
                after = self._colour_library_cell_state(ws.cell(row, column))
                if before == after:
                    continue
                name_key = (self._colour_name_key(str(ws.cell(row, english_column).value or ''))
                            if row > 1 else None)
                written.append((row, column, before, after, name_key))

            if written:
                self._save_colour_library(book, path, original_stat)
                return path, sheet, written
        finally:
            book.close()

    def _undo_colour_library_write(self, path, sheet, written):
        original_stat = os.stat(path)
        book = load_workbook(path, keep_vba=path.lower().endswith('.xlsm'))
        try:
            ws = book[sheet]
            english_column = next(cell.column for cell in ws[1]
                                  if str(cell.value).strip() == '英文名称')
            names = {row: self._colour_name_key(str(ws.cell(row, english_column).value or ''))
                     for row, *_ in written if row > 1}
            changed = False
            # Undo data first so that newly created, empty column headers can be restored too.
            for row, column, before, imported, name_key in sorted(
                    written, key=lambda change: change[0] == 1):
                cell = ws.cell(row, column)
                current = self._colour_library_cell_state(cell)
                if current[:2] != imported[:2]:
                    continue
                if name_key is not None and names.get(row) != name_key:
                    continue
                if (row == 1 and before[:2] != imported[:2] and any(
                        ws.cell(data_row, column).value not in (None, '')
                        for data_row in range(2, ws.max_row + 1))):
                    continue
                # Keep later edits to formatting, notes or links on this cell.
                metadata = tuple(old if now == saved else now
                                 for old, now, saved in zip(before[2:], current[2:], imported[2:]))
                self._restore_colour_library_cell(cell, before[:2] + metadata)
                changed = True
            if changed:
                self._save_colour_library(book, path, original_stat)
        finally:
            book.close()

    def _populate_dests(self):
        for item in self.dest_tv.get_children():
            self.dest_tv.delete(item)
        if self.df is None or 'DESTINATION NAME' not in self.df.columns:
            return
        seen = []
        for dest in self.df['DESTINATION NAME']:
            if dest not in seen:
                seen.append(dest)
        for dest in seen:
            code = dest.split(' - ')[0].strip()
            port_cn, port_code = self._lookup_port(dest)
            self.dest_tv.insert('', 'end', values=(code, dest, port_cn, port_code))

    # ── Clipboard paste into colour table ─────────────────────────────────
    def _paste_colour_table(self, event=None):
        try:
            raw = self.clipboard_get()
        except Exception:
            return
        lines = raw.strip().splitlines()
        if not lines:
            return
        editing_cell = event is not None and isinstance(event.widget, tk.Entry)
        if editing_cell and len(lines) == 1 and '\t' not in raw:
            return  # 单个值交给输入框处理，Excel 多格数据仍按表格粘贴

        self.colour_tv.finish_edit()
        if editing_cell:
            self.colour_tv.focus_set()
        all_items = self.colour_tv.get_children()
        if not all_items:
            return

        selected = self.colour_tv.selection()
        start_idx = all_items.index(selected[0]) if selected else 0

        for i, line in enumerate(lines):
            idx = start_idx + i
            if idx >= len(all_items):
                break
            item = all_items[idx]
            fields = line.rstrip('\r').split('\t')
            vals = list(self.colour_tv.item(item, 'values'))
            # Map pasted columns → 颜色编号(2), 颜色名称(3), 英文颜色(4), 厂款号(6), 工厂出厂期(9)
            paste_targets = (2, 3, 4, 6, 9)
            for j, field in enumerate(fields[:5]):
                vals[paste_targets[j]] = field.strip()
            self.colour_tv.item(item, values=vals)

        self._refresh_preview()
        return 'break'

    # ── 多拼色拆分弹窗 ─────────────────────────────────────────────────────
    def _open_colour_split(self, event):
        """双击「客户色号(0)/客户色名(1)」只读列 → 打开多拼色拆分弹窗。"""
        if self.colour_tv.identify('region', event.x, event.y) != 'cell':
            return
        col_id = self.colour_tv.identify_column(event.x)
        try:
            col_idx = int(col_id.replace('#', '')) - 1
        except ValueError:
            return
        if col_idx not in (0, 1):
            return                                      # 其余列交给基类内编辑
        row_id = self.colour_tv.identify_row(event.y)
        if not row_id:
            return
        self._colour_split_dialog(row_id)
        return 'break'

    @staticmethod
    def _seed_colour_parts(raw, cur_cn, prior_en_split):
        """返回 (cn_parts, en_parts)，长度一致；作为弹窗各段初值。

        英文种子：若曾拆分过（prior_en_split 非空，'|' 连接）→ 按 '|' 恢复各段；
        否则对原始英文**优先按逗号**、无逗号再按 '/'、都没有则单段
        （内含斜杠的名称如 'C/D'、'P/P' 在有逗号时被保留不拆）。
        中文沿用已填的 '/' 分段（默认空，按英文段数补齐）。
        """
        prior = str(prior_en_split or '').strip()
        if prior:
            en_parts = [p.strip() for p in prior.split('|') if p.strip()]
        else:
            en_src = str(raw or '').strip()
            if ',' in en_src:
                en_parts = [p.strip() for p in en_src.split(',') if p.strip()]
            elif '/' in en_src:
                en_parts = [p.strip() for p in en_src.split('/') if p.strip()]
            else:
                en_parts = [en_src]
        if not en_parts:
            en_parts = ['']
        cur_cn = str(cur_cn or '').strip()
        cn_parts = [p.strip() for p in cur_cn.split('/')] if cur_cn else []
        n = max(len(en_parts), 1)
        en_parts = (en_parts + [''] * n)[:n]
        cn_parts = (cn_parts + [''] * n)[:n]            # 中文不足按英文段数补空
        return cn_parts, en_parts

    @staticmethod
    def _seed_fac_parts(default_fac, prior_fac_split, n):
        """返回长度 n 的逐件厂款号种子：曾拆分过则按 '|' 恢复；否则整行现有厂款号重复填满
        （用户只需改动不同的那一件，如 3 件中 2 件同厂款号、1 件不同）。"""
        prior = str(prior_fac_split or '').strip()
        if prior:
            parts = [p.strip() for p in prior.split('|')]
        else:
            parts = [str(default_fac or '').strip()] * n
        parts = (parts + [parts[-1] if parts else ''] * n)[:n]
        return parts

    def _colour_split_dialog(self, row_id):
        vals = list(self.colour_tv.item(row_id, 'values'))
        cust_code = vals[0]
        raw       = vals[1]
        cn_seed, en_seed = self._seed_colour_parts(
            raw, vals[3], self._en_split.get(cust_code, ''))
        fac_seed = self._seed_fac_parts(
            vals[6], self._fac_split.get(cust_code, ''), len(en_seed))

        win = tk.Toplevel(self)
        win.title(f'拆分多拼色 — {cust_code}')
        win.transient(self.winfo_toplevel())
        win.grab_set()
        win.resizable(False, False)

        ttk.Label(win, text=f'客户色号：{cust_code}',
                  font=('Arial', 10, 'bold')).grid(
            row=0, column=0, columnspan=3, sticky='w', padx=10, pady=(10, 0))
        ttk.Label(win, text=f'原始英文：{raw}', foreground='gray',
                  wraplength=460, justify='left').grid(
            row=1, column=0, columnspan=3, sticky='w', padx=10, pady=(2, 8))

        topbar = ttk.Frame(win)
        topbar.grid(row=2, column=0, columnspan=3, sticky='w', padx=10)
        ttk.Label(topbar, text='颜色件数 N：').pack(side='left')
        n_var = tk.IntVar(value=max(len(en_seed), 1))
        spin = ttk.Spinbox(topbar, from_=1, to=8, width=5, textvariable=n_var)
        spin.pack(side='left')
        ttk.Label(topbar, text='（每段：中文颜色 + 英文颜色 + 厂款号；厂款号默认沿用整行现有值，'
                              '只需改动不同的那一件）',
                  foreground='gray').pack(side='left', padx=(8, 0))

        rows_frame = ttk.Frame(win)
        rows_frame.grid(row=3, column=0, columnspan=3, sticky='ew', padx=10, pady=8)
        ttk.Label(rows_frame, text='#', width=3).grid(row=0, column=0)
        ttk.Label(rows_frame, text='中文 颜色名称', width=22).grid(row=0, column=1)
        ttk.Label(rows_frame, text='英文颜色', width=30).grid(row=0, column=2)
        ttk.Label(rows_frame, text='厂款号', width=16).grid(row=0, column=3)

        cn_seed = list(cn_seed)
        en_seed = list(en_seed)
        fac_seed = list(fac_seed)
        entries = []                                    # [(cn_entry, en_entry, fac_entry), ...]

        def _rebuild_rows():
            try:
                n = int(n_var.get())
            except (tk.TclError, ValueError):
                return
            n = max(1, min(8, n))
            # 保留已输入值
            cur_cn = [e[0].get() for e in entries]
            cur_en = [e[1].get() for e in entries]
            cur_fac = [e[2].get() for e in entries]
            for e in entries:
                e[0].destroy()
                e[1].destroy()
                e[2].destroy()
            for lbl in getattr(_rebuild_rows, '_labels', []):
                lbl.destroy()
            entries.clear()
            labels = []
            for i in range(n):
                lbl = ttk.Label(rows_frame, text=str(i + 1), width=3)
                lbl.grid(row=i + 1, column=0, pady=1)
                labels.append(lbl)
                cn_e = ttk.Entry(rows_frame, width=24)
                cn_e.grid(row=i + 1, column=1, padx=2, pady=1)
                en_e = ttk.Entry(rows_frame, width=32)
                en_e.grid(row=i + 1, column=2, padx=2, pady=1)
                fac_e = ttk.Entry(rows_frame, width=18)
                fac_e.grid(row=i + 1, column=3, padx=2, pady=1)
                cn_val = cur_cn[i] if i < len(cur_cn) else (
                    cn_seed[i] if i < len(cn_seed) else '')
                en_val = cur_en[i] if i < len(cur_en) else (
                    en_seed[i] if i < len(en_seed) else '')
                fac_val = cur_fac[i] if i < len(cur_fac) else (
                    fac_seed[i] if i < len(fac_seed) else (fac_seed[-1] if fac_seed else ''))
                cn_e.insert(0, cn_val)
                en_e.insert(0, en_val)
                fac_e.insert(0, fac_val)
                entries.append((cn_e, en_e, fac_e))
            _rebuild_rows._labels = labels

        _rebuild_rows()
        # trace 覆盖 Spinbox 箭头与键入；clamp 在 _rebuild_rows 内处理
        n_var.trace_add('write', lambda *_: _rebuild_rows())

        def _confirm():
            cn  = '/'.join(e[0].get().strip() for e in entries)    # 中文用 /
            en  = '|'.join(e[1].get().strip() for e in entries)    # 英文用 | (段内 / 得以保留)
            fac = '|'.join(e[2].get().strip() for e in entries)    # 厂款号用 | (逐件对齐)
            if any(not e[0].get().strip() or not e[1].get().strip()
                   for e in entries):
                if not messagebox.askyesno(
                        '存在空白', '有颜色段的中文或英文为空，仍要应用吗？',
                        parent=win):
                    return
            new_vals = list(self.colour_tv.item(row_id, 'values'))
            new_vals[3] = cn                            # 写「颜色名称」(中文, / 连接)
            # 「英文颜色」(col 4)、「厂款号」(col 6) 保留导入原文/整行现有值，不覆盖；
            # 逐件英文/厂款号另存入末列（供步骤②按件对齐读取）
            self._en_split[new_vals[0]] = en
            self._fac_split[new_vals[0]] = fac
            self.colour_tv.item(row_id, values=new_vals)
            self.colour_tv.event_generate('<<RowEdited>>')
            win.destroy()

        btns = ttk.Frame(win)
        btns.grid(row=4, column=0, columnspan=3, sticky='e', padx=10, pady=(4, 10))
        ttk.Button(btns, text='确定', command=_confirm).pack(side='right', padx=4)
        ttk.Button(btns, text='取消', command=win.destroy).pack(side='right')

    # ── Helpers ───────────────────────────────────────────────────────────
    def _colour_map(self):
        """Returns {cust_code: (colour_id, colour_cn, colour_en, fac_style, fac_style2, fac_date)}"""
        result = {}
        for item in self.colour_tv.get_children():
            vals = self.colour_tv.item(item, 'values')
            cc, cn, cid, ccn, cen = vals[:5]
            fac_style  = vals[6] if len(vals) > 6 else ''
            fac_style2 = vals[8] if len(vals) > 8 else ''
            fac_date   = vals[9] if len(vals) > 9 else ''
            result[cc] = (cid, ccn, cen, fac_style, fac_style2, fac_date)
        return result

    def _dest_map(self):
        result = {}
        for item in self.dest_tv.get_children():
            _, full, port_cn, port_code = self.dest_tv.item(item, 'values')
            result[full] = (port_cn, port_code)
        return result

    @staticmethod
    def _parse_date(raw):
        for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%m/%d/%Y', '%Y/%m/%d'):
            try:
                return datetime.strptime(str(raw).strip(), fmt).strftime('%Y-%m-%d')
            except ValueError:
                pass
        v = str(raw).strip()
        return '' if v in ('', 'nan', 'None', 'NaT') else v

    # ── Build output rows ─────────────────────────────────────────────────
    def _build_rows(self, style_codes=None):
        """Build rows. If style_codes given, only include those codes."""
        if self.df is None:
            return []
        dm = self._dest_map()
        cm = self._colour_map()

        freight_raw = (self.df['FREIGHT MODE'].iloc[0]
                       if 'FREIGHT MODE' in self.df.columns else 'SEA')
        freight = FREIGHT_MAP.get(str(freight_raw).upper().strip(), str(freight_raw))

        seen_codes, colour_pairs = [], []
        for code, en in zip(
                self.df.get('STYLE COLOUR CODE',
                            pd.Series([''] * len(self.df))),
                self.df.get('COLOUR DESCRIPTION',
                            pd.Series([''] * len(self.df)))):
            code = str(code)
            if code not in seen_codes:
                seen_codes.append(code)
                colour_pairs.append((code, str(en)))

        if style_codes is not None:
            colour_pairs = [(c, e) for c, e in colour_pairs if c in style_codes]

        dests = (list(dict.fromkeys(self.df['DESTINATION NAME'].tolist()))
                 if 'DESTINATION NAME' in self.df.columns else [])

        rows = []
        sn = 0
        for cust_code, cust_name in colour_pairs:
            colour_id, colour_cn, colour_en, fac_style, fac_style2, fac_date = cm.get(
                cust_code, ('', '', '', '', '', ''))
            colour_df = (self.df[self.df['STYLE COLOUR CODE'] == cust_code]
                         if 'STYLE COLOUR CODE' in self.df.columns else self.df)

            for dest in dests:
                port_cn, port_code = dm.get(dest, ('', ''))
                sub = (colour_df[colour_df['DESTINATION NAME'] == dest]
                       if 'DESTINATION NAME' in colour_df.columns
                       else pd.DataFrame())
                if sub.empty:
                    continue

                handover  = self._parse_date(
                    sub['FINAL HANDOVER DATE'].iloc[0]
                    if 'FINAL HANDOVER DATE' in sub.columns else '')
                ship_date = self._parse_date(
                    sub['LATEST SHIP-BY DATE'].iloc[0]
                    if 'LATEST SHIP-BY DATE' in sub.columns else '')

                qty_map = {str(r['SIZE']).strip(): int(r['QUANTITY'])
                           for _, r in sub.iterrows()}
                total = sum(qty_map.values())

                sn += 1
                rows.append({
                    '颜色编号':            colour_id,
                    '颜色名称':            colour_cn,
                    '英文颜色':            colour_en,
                    '走货方式':            freight,
                    '流水号':              sn,
                    '目的港':              port_cn,
                    '港口代码':            port_code,
                    '实际出厂（年-月-日）': handover,
                    '客户船期（年-月-日）': ship_date,
                    '价格':                None,
                    '客户色号':            cust_code,
                    '客户色名':            colour_cn,
                    '工厂出厂期':          fac_date.strip() or None,
                    '厂款号':              fac_style.strip() or None,
                    '厂款号2':             fac_style2.strip() or None,
                    **{s: qty_map.get(s) for s in self.sizes},
                    '数量合计':            total,
                    '客款号':              cust_code,
                    '交期1':               None,   # intentionally empty
                    # 逐件英文/厂款号（'|' 连接）；空=未拆分。写到输出最末列供步骤②读取
                    '英文颜色拆分':        self._en_split.get(cust_code, ''),
                    '厂款号拆分':          self._fac_split.get(cust_code, ''),
                })
        return rows

    def _refresh_preview(self):
        if self.prev_tv is None:
            return
        for item in self.prev_tv.get_children():
            self.prev_tv.delete(item)
        if self.df is None:
            if hasattr(self, 'fname_note'):
                self.fname_note.set('')
            return
        if hasattr(self, 'fname_note'):
            names = [n + '.xlsx' for n, _ in self._planned_outputs()]
            self.fname_note.set('导出文件：' + ' ， '.join(names) if names else '')
        cols = getattr(self, '_prev_cols', None) or []
        for base in self._build_rows():
            disp_rows = (self._expand_preview_row(base)
                         if self._preview_split else [base])
            for row in disp_rows:
                vals = [self._preview_cell(row, col) for col in cols]
                self.prev_tv.insert('', 'end', values=vals)

    def _preview_cell(self, row, col):
        """按预览列名取行值（尺码列直接取；差异列名走 _PREVIEW_KEY 映射）。"""
        if col in self.sizes:
            return str(row.get(col) or '')
        key = self._PREVIEW_KEY.get(col, col)
        v = row.get(key, '')
        return '' if v is None else v

    def _expand_preview_row(self, row):
        """逐件拆分视图：多拼色一行展开为一行一件（中文按 / 、英文/厂款号按 |）。"""
        cn = str(row.get('颜色名称') or '')
        ensplit = str(row.get('英文颜色拆分') or '')
        facsplit = str(row.get('厂款号拆分') or '')
        cn_parts = [p.strip() for p in cn.split('/') if p.strip()]
        en_parts = [p.strip() for p in ensplit.split('|') if p.strip()] if ensplit else []
        fac_parts = [p.strip() for p in facsplit.split('|') if p.strip()] if facsplit else []
        if len(cn_parts) <= 1 and len(en_parts) <= 1 and len(fac_parts) <= 1:
            return [row]
        n = max(len(cn_parts), len(en_parts), len(fac_parts), 1)
        out = []
        for i in range(n):
            r = dict(row)
            if i < len(cn_parts):
                r['颜色名称'] = cn_parts[i]
            if i < len(en_parts):
                r['英文颜色'] = en_parts[i]
            if i < len(fac_parts):
                r['厂款号'] = fac_parts[i]
            out.append(r)
        return out

    # ── Write to Excel ────────────────────────────────────────────────────
    @staticmethod
    def _copy_style(src, dst):
        if src.has_style:
            dst.font          = copy(src.font)
            dst.border        = copy(src.border)
            dst.fill          = copy(src.fill)
            dst.number_format = src.number_format
            dst.alignment     = copy(src.alignment)

    def _write_excel(self, template_path, rows, save_path):
        wb = load_workbook(template_path)
        if 'OutSheet' not in wb.sheetnames:
            raise ValueError("模板中未找到名为 'OutSheet' 的工作表")
        ws = wb['OutSheet']

        # ── Locate size zone in template header ──────────────────────────
        hdr_vals = [ws.cell(row=1, column=c).value
                    for c in range(1, ws.max_column + 1)]
        hdr_str  = [str(v).strip() if v is not None else '' for v in hdr_vals]

        # Size columns = any header not in KNOWN_FIXED_COLS (and non-empty)
        size_idxs = [i for i, h in enumerate(hdr_str)
                     if h and h not in KNOWN_FIXED_COLS]

        if size_idxs:
            zone_start_col = size_idxs[0] + 1   # 1-indexed
            current_count  = len(size_idxs)
        else:
            # No size columns in template yet — insert before first POST_SIZE col
            fallback = next(
                (i for i, h in enumerate(hdr_str) if h in POST_SIZE_COLS),
                len(hdr_str))
            zone_start_col = fallback + 1        # 1-indexed (insert before post-zone)
            current_count  = 0

        needed = len(self.sizes)
        diff   = needed - current_count

        if diff > 0:
            # Add columns at end of current size zone
            insert_at = zone_start_col + current_count
            ws.insert_cols(insert_at, diff)
            # Copy column width from the last size column
            if current_count > 0:
                ref_letter = get_column_letter(zone_start_col + current_count - 1)
                ref_width  = ws.column_dimensions[ref_letter].width or 8
                for k in range(diff):
                    ltr = get_column_letter(insert_at + k)
                    ws.column_dimensions[ltr].width = ref_width
        elif diff < 0:
            # Remove excess size columns from the end of the zone
            ws.delete_cols(zone_start_col + needed, -diff)

        # Write size headers (copy style from first existing size col if available)
        ref_hdr_cell = ws.cell(row=1, column=zone_start_col) if current_count > 0 else None
        for j, size in enumerate(self.sizes):
            cell = ws.cell(row=1, column=zone_start_col + j)
            cell.value = size
            if ref_hdr_cell and ref_hdr_cell.has_style:
                self._copy_style(ref_hdr_cell, cell)

        # ── Rebuild col_map from updated header ───────────────────────────
        def _build_col_map():
            hdr = [ws.cell(row=1, column=c).value
                   for c in range(1, ws.max_column + 1)]
            return {str(v).strip(): i + 1
                    for i, v in enumerate(hdr) if v is not None}

        col_map = _build_col_map()

        # ── Ensure 厂款号2 column exists when needed ──────────────────────
        if ('厂款号2' not in col_map
                and any(r.get('厂款号2') for r in rows)):
            anchor = col_map.get('厂款号')
            if anchor:
                pos = anchor + 1
                ws.insert_cols(pos)
                ltr_anchor = get_column_letter(anchor)
                ws.column_dimensions[get_column_letter(pos)].width = \
                    ws.column_dimensions[ltr_anchor].width or 10
                hdr_cell = ws.cell(row=1, column=pos, value='厂款号2')
                self._copy_style(ws.cell(row=1, column=anchor), hdr_cell)
            else:
                pos = ws.max_column + 1
                ws.cell(row=1, column=pos, value='厂款号2')
            col_map = _build_col_map()

        # ── Ensure 英文颜色拆分 column exists at the very end when needed ───
        if ('英文颜色拆分' not in col_map
                and any(r.get('英文颜色拆分') for r in rows)):
            pos = ws.max_column + 1
            hdr_cell = ws.cell(row=1, column=pos, value='英文颜色拆分')
            anchor = col_map.get('英文颜色')
            if anchor:
                self._copy_style(ws.cell(row=1, column=anchor), hdr_cell)
                ws.column_dimensions[get_column_letter(pos)].width = \
                    ws.column_dimensions[get_column_letter(anchor)].width or 14
            col_map = _build_col_map()

        # ── Ensure 厂款号拆分 column exists at the very end when needed ───
        if ('厂款号拆分' not in col_map
                and any(r.get('厂款号拆分') for r in rows)):
            pos = ws.max_column + 1
            hdr_cell = ws.cell(row=1, column=pos, value='厂款号拆分')
            anchor = col_map.get('厂款号')
            if anchor:
                self._copy_style(ws.cell(row=1, column=anchor), hdr_cell)
                ws.column_dimensions[get_column_letter(pos)].width = \
                    ws.column_dimensions[get_column_letter(anchor)].width or 14
            col_map = _build_col_map()

        # ── Append data rows ──────────────────────────────────────────────
        last = ws.max_row
        for i, row_data in enumerate(rows):
            new_row = last + i + 1
            if last >= 2:
                for col in range(1, ws.max_column + 1):
                    self._copy_style(ws.cell(row=last, column=col),
                                     ws.cell(row=new_row, column=col))
            for col_name, col_idx in col_map.items():
                val = row_data.get(col_name)
                if val is not None:
                    ws.cell(row=new_row, column=col_idx, value=val)

        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        _set_output_font(cell)
        wb.save(save_path)

    # ── Generate ──────────────────────────────────────────────────────────
    def _generate(self):
        if self._batch_pos:                       # 批量模式：按 PO 各自出文件
            return self._batch_write()
        xlsx_path = self.xlsx_var.get().strip()
        if not xlsx_path or not os.path.exists(xlsx_path):
            messagebox.showerror('错误', '请选择有效的订单明细模板'); return
        if self.df is None:
            messagebox.showerror('错误', '请先加载 PO 数据'); return

        # Filenames come from each file's actual 厂款号/厂款号2 (step ②),
        # written into a user-chosen folder.
        outputs = self._planned_outputs()
        if not outputs:
            messagebox.showerror('错误', '没有可填充的数据'); return

        out_dir = self.out_dir_var.get().strip()
        if not out_dir:
            messagebox.showerror('错误', '请在 ① 文件选择 中指定导出文件夹'); return
        if not os.path.isdir(out_dir):
            try:
                os.makedirs(out_dir, exist_ok=True)
            except Exception as ex:
                messagebox.showerror('错误', f'无法创建导出文件夹：{ex}'); return

        self._save_config()
        saved = []
        for name, rows in outputs:
            save_path = os.path.join(out_dir, f'{name}.xlsx')
            try:
                self._write_excel(xlsx_path, rows, save_path)
                saved.append(save_path)
            except Exception as ex:
                messagebox.showerror('错误', f'{name}.xlsx：{ex}'); return

        msg = f'已生成 {len(saved)} 个文件：\n'
        for path in saved:
            msg += f'  {os.path.basename(path)}\n'
        msg += f'\n保存至：\n{out_dir}'
        messagebox.showinfo('完成 ✓', msg)
        self.status.set(f'已生成 {len(saved)} 个文件 → {out_dir}')

    # ── Batch：扫描文件夹 → 合并颜色表填写 → 一次生成全部 ────────────────
    def _batch_scan(self):
        """选一个 PO 文件夹，解析全部 PO，把所有客户色号/目的地**合并去重**后
        填进现有的颜色表 / 目的地表，供用户像单个 PO 那样填写一次，
        然后在 ③ 页点「生成 订单明细」一次性写出全部文件。"""
        src_dir = filedialog.askdirectory(title='选择存放 PO 文件的文件夹')
        if not src_dir:
            return

        # 同名 pdf/csv 只取一个：先放 csv，再 setdefault 放 pdf → CSV 优先
        pairs = {}
        for ext in ('.csv', '.pdf'):
            for fn in sorted(os.listdir(src_dir)):
                if fn.startswith('~$') or not fn.lower().endswith(ext):
                    continue
                pairs.setdefault(os.path.splitext(fn)[0],
                                 os.path.join(src_dir, fn))
        if not pairs:
            messagebox.showwarning('提示', '该文件夹中没有 .pdf / .csv 文件')
            return

        pos, dfs, skipped = {}, [], []
        pack_colour_codes = set()
        for stem in sorted(pairs):
            try:
                path = pairs[stem]
                df = load_po(path)
                if path.lower().endswith('.csv'):
                    pack_colour_codes.update(self._csv_pack_colour_codes(df))
                for col in ('SIZE', 'STYLE COLOUR CODE', 'DESTINATION NAME'):
                    if col not in df.columns:
                        raise ValueError('缺少列 ' + col)
                po_no = (str(df['ORDER NUMBER'].iloc[0]).strip()
                         if 'ORDER NUMBER' in df.columns else stem)
                pos[stem] = (po_no, df)
                dfs.append(df)
            except Exception as ex:                    # noqa: BLE001
                skipped.append((stem, f'{type(ex).__name__}: {ex}'))
        if not pos:
            messagebox.showerror(
                '批量扫描失败',
                '没有可用的 PO：\n' + _clip_list([f'{n}：{r}' for n, r in skipped]))
            return

        # 合并成一张总表：用于填颜色/目的地表与预览（生成时仍按 PO 各自出文件）
        self._batch_pos = pos
        self._batch_dir = src_dir
        self._pack_colour_codes = pack_colour_codes
        self.df = pd.concat(dfs, ignore_index=True)
        self.sizes = sorted(
            self.df['SIZE'].astype(str).str.strip().unique().tolist(),
            key=_size_sort_key)
        self.po_number = ''

        self._load_port_data()
        self._populate_colours()
        self._populate_dests()

        # 厂款号：按客款号从下单价格表预填（用户仍可改）
        fac_map, fac_conflicts, price_err = {}, [], ''
        price = self.price_var.get().strip()
        if price:
            try:
                fac_map, fac_conflicts = load_fac_style_map(price)
            except Exception as ex:                    # noqa: BLE001
                price_err = str(ex)
        miss_style = []
        for iid in self.colour_tv.get_children():
            vals = list(self.colour_tv.item(iid, 'values'))
            style = str(vals[0]).split('-')[0].strip()
            fac = fac_map.get(style, '')
            if fac:
                vals[6] = fac
                self.colour_tv.item(iid, values=vals)
            elif style not in miss_style:
                miss_style.append(style)

        miss_port = [str(self.dest_tv.item(i, 'values')[1])
                     for i in self.dest_tv.get_children()
                     if not str(self.dest_tv.item(i, 'values')[2]).strip()]

        self._rebuild_preview_tv()
        self._save_config()
        n_col = len(self.colour_tv.get_children())
        self.gen_note.set(
            f'批量模式：{len(pos)} 个 PO 待生成（填好后在 ③ 页点「生成 订单明细」）')
        self.status.set(
            f'批量已扫描 {len(pos)} 个 PO · {n_col} 个客户色号 · '
            f'尺码：{", ".join(self.sizes)}')

        msg = (f'已扫描 {len(pos)} 个 PO，合并出 {n_col} 个客户色号。\n\n'
               '请在 ② 字段配置 核对颜色编号和颜色名称，并填写工厂出厂期'
               '（厂款号已按价格表预填），\n'
               '填好后到 ③ 预览 & 生成 点「生成 订单明细」，一次写出全部文件。')
        if price_err:
            msg += '\n\n⚠ 价格表读取失败，厂款号未预填：\n  ' + price_err
        elif not price:
            msg += '\n\n⚠ 未选择「下单价格表」，厂款号未预填。'
        if miss_style:
            msg += ('\n\n⚠ 价格表中找不到以下客款号（厂款号需手填）：\n'
                    + _clip_list(miss_style))
        if fac_conflicts:
            msg += ('\n\n⚠ 价格表中以下客款号有多个不同厂款号（取第一个）：\n'
                    + _clip_list(fac_conflicts))
        if miss_port:
            msg += ('\n\n⚠ 港口文件中找不到以下目的地（需手填目的港/港口代码）：\n'
                    + _clip_list(miss_port))
        if skipped:
            msg += (f'\n\n⚠ 跳过 {len(skipped)} 个文件：\n'
                    + _clip_list([f'{n}：{r}' for n, r in skipped]))
        messagebox.showinfo('批量扫描完成', msg)
        self.nb.select(1)

    def _batch_write(self):
        """批量模式的生成：颜色表/目的地表是全文件夹共用的，按 PO 各自出文件。"""
        xlsx_path = self.xlsx_var.get().strip()
        if not xlsx_path or not os.path.exists(xlsx_path):
            messagebox.showerror('错误', '请选择有效的订单明细模板')
            return
        out_dir = self.out_dir_var.get().strip()
        if not out_dir:
            messagebox.showerror('错误', '请在 ① 文件选择 中指定导出文件夹')
            return
        if not os.path.isdir(out_dir):
            try:
                os.makedirs(out_dir, exist_ok=True)
            except Exception as ex:                    # noqa: BLE001
                messagebox.showerror('错误', f'无法创建导出文件夹：{ex}')
                return

        cm, dm = self._colour_map(), self._dest_map()   # 全文件夹共用（用户已填）
        en_split, fac_split = dict(self._en_split), dict(self._fac_split)
        self._save_config()

        saved, skipped, used = [], [], set()
        for stem, (po_no, df) in self._batch_pos.items():
            try:
                a = App.__new__(App)
                a.df = df
                a.sizes = sorted(
                    df['SIZE'].astype(str).str.strip().unique().tolist(),
                    key=_size_sort_key)
                a.po_number = po_no
                a._en_split, a._fac_split = en_split, fac_split
                a._colour_map = lambda cm=cm: cm
                a._dest_map = lambda dm=dm: dm

                outputs = a._planned_outputs()
                if not outputs:
                    skipped.append((stem, '没有可填充的数据'))
                    continue
                for name, rows in outputs:
                    # 跨 PO / 重跑 去重（_planned_outputs 的 used 只在单次调用内有效）
                    final, k = name, 2
                    while (final in used
                           or os.path.exists(os.path.join(out_dir, f'{final}.xlsx'))):
                        final, k = f'{name}_{k}', k + 1
                    used.add(final)
                    a._write_excel(xlsx_path, rows,
                                   os.path.join(out_dir, f'{final}.xlsx'))
                    saved.append(final)
            except Exception as ex:                    # noqa: BLE001 — 单个失败不中断
                skipped.append((stem, f'{type(ex).__name__}: {ex}'))

        msg = (f'已生成 {len(saved)} 个文件（{len(self._batch_pos)} 个 PO）\n'
               f'保存至：\n{out_dir}')
        if skipped:
            msg += (f'\n\n⚠ 跳过 {len(skipped)} 个：\n'
                    + _clip_list([f'{n}：{r}' for n, r in skipped]))
        messagebox.showinfo('批量生成完成 ✓', msg)
        self.status.set(f'批量生成完成：{len(saved)} 个文件 → {out_dir}')


if __name__ == '__main__':
    root = tk.Tk()
    root.title('PO → 订单明细 导入工具  v1.6')
    root.geometry('1150x860')
    root.minsize(900, 700)
    App(root)
    root.mainloop()
