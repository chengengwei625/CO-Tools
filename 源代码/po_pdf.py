# -*- coding: utf-8 -*-
"""PO 文件读取（PDF / CSV）+ 下单价格表查表。

功能① 原本只吃 PO 的 .csv。客户同时提供同名 .pdf，本模块把 PDF 解析成
与 CSV **完全相同列名/形状** 的 DataFrame（长表：每 style×destination×size 一行），
因此下游 po_import_app / generate_tabs 全部无需改动。

PDF 为文字型（非扫描），用 PyMuPDF(fitz) 取文本行即可，无需 OCR / 坐标。
"""
from __future__ import annotations

import re

import pandas as pd

# CSV 的列名即契约；PDF 解析结果必须一致
PO_COLUMNS = [
    'ORDER NUMBER', 'DEVELOPMENT CODE', 'STYLE COLOUR CODE', 'STYLE DESCRIPTION',
    'COLOUR DESCRIPTION', 'DESTINATION NAME', 'SIZE', 'QUANTITY',
    'FINAL BOOKED BY DATE', 'FINAL HANDOVER DATE', 'LATEST SHIP-BY DATE',
    'FREIGHT MODE',
]

# 目的地块起始行：形如 "8094 - BASH (WHOLESALE) - FOSCHINI ..."
_DEST_START_RE = re.compile(r'^\d{3,5}\s+-\s+')
_STYLE_HDR = 'DEVELOPMENT CODE'
# 款式子块的表头标签（顶部汇总表还带 FOB 两列）
_STYLE_LABELS = frozenset((
    'DEVELOPMENT CODE', 'STYLE COLOUR CODE', 'STYLE DESCRIPTION',
    'COLOUR DESCRIPTION', 'FOB CURRENCY', 'FOB VALUE',
))


def _lines(path) -> list[str]:
    """全部页面展平成非空文本行。fitz 默认顺序已把尺码/数量交替排好。"""
    import fitz                                       # PyMuPDF；打包见 .spec

    with fitz.open(str(path)) as doc:
        out = []
        for page in doc:
            out.extend(l.strip() for l in page.get_text().split('\n') if l.strip())
    return out


def _after(lines: list[str], label: str, default: str = '') -> str:
    """取 label 行的下一行（PDF 的 key/value 是上下两行）。"""
    try:
        return lines[lines.index(label) + 1].strip()
    except (ValueError, IndexError):
        return default


def _value_in(block: list[str], label: str, default: str = '') -> str:
    for i, l in enumerate(block):
        if l == label and i + 1 < len(block):
            return block[i + 1].strip()
    return default


def _dest_name(block: list[str]) -> str:
    """目的地名 = 块起点起、到第一个含逗号的行之前（地址行必含 ','，名称行不含）。"""
    parts = []
    for l in block:
        if ',' in l:
            break
        parts.append(l)
    return re.sub(r'\s+', ' ', ' '.join(parts)).strip()


def _styles_in(block: list[str]) -> list[dict]:
    """解析目的地块内的各款式子块。

    版式（实测）：先若干个表头标签，再等量的值，然后尺码/数量**交替**，最后 TOTAL：
        DEVELOPMENT CODE / STYLE COLOUR CODE / STYLE DESCRIPTION / COLOUR DESCRIPTION
        D3012764 / 6339064-01 / REMI LACE ... / BLACK
        2XS / 5 / XS / 18 / ... / TOTAL / 120
    标签数量按实际出现的取（顶部汇总表还带 FOB CURRENCY / FOB VALUE）。
    """
    idxs = [i for i, l in enumerate(block) if l == _STYLE_HDR]
    styles = []
    for n, start in enumerate(idxs):
        end = idxs[n + 1] if n + 1 < len(idxs) else len(block)
        seg = block[start:end]

        # 1) 连续的表头标签 → 2) 等量的值
        labels = []
        j = 0
        while j < len(seg) and seg[j] in _STYLE_LABELS:
            labels.append(seg[j])
            j += 1
        values = seg[j:j + len(labels)]
        if len(values) < len(labels):
            continue
        f = dict(zip(labels, (v.strip() for v in values)))
        scc = f.get('STYLE COLOUR CODE', '')

        # 3) 尺码/数量交替，直到 TOTAL
        pairs, total = [], None
        rest = seg[j + len(labels):]
        for i, l in enumerate(rest):
            if l == 'TOTAL':
                total = rest[i + 1].strip() if i + 1 < len(rest) else None
                break
            pairs.append(l)
        if total is None:
            continue
        if len(pairs) % 2:
            raise ValueError('尺码/数量行数不成对：%s / %s' % (scc, pairs))

        sizes, qtys = pairs[0::2], pairs[1::2]
        if not all(_is_int(q) for q in qtys):
            raise ValueError('数量非数字：%s / %s' % (scc, qtys))
        qtys_i = [int(float(q)) for q in qtys]
        if sum(qtys_i) != int(float(total)):
            raise ValueError('数量合计与 TOTAL 不符：%s (%d != %s)'
                             % (scc, sum(qtys_i), total))

        styles.append({'dev': f.get('DEVELOPMENT CODE', ''), 'scc': scc,
                       'sdesc': f.get('STYLE DESCRIPTION', ''),
                       'cdesc': f.get('COLOUR DESCRIPTION', ''),
                       'sizes': sizes, 'qtys': qtys_i})
    return styles


def _is_int(s: str) -> bool:
    try:
        int(float(s))
        return True
    except (TypeError, ValueError):
        return False


def parse_po_pdf(path) -> pd.DataFrame:
    """把一个 PO PDF 解析成与 CSV 同形的长表 DataFrame。"""
    lines = _lines(path)
    order_no = _after(lines, 'ORDER NUMBER')
    freight = _after(lines, 'FREIGHT MODE')

    starts = [i for i, l in enumerate(lines) if _DEST_START_RE.match(l)]
    rows = []
    for n, s in enumerate(starts):
        e = starts[n + 1] if n + 1 < len(starts) else len(lines)
        block = lines[s:e]
        booked = _value_in(block, 'FINAL BOOKED BY DATE')
        if not booked:
            # 第 1 页顶部的款式汇总表：没有日期，数量与后面重复，跳过
            continue
        dest = _dest_name(block)
        handover = _value_in(block, 'FINAL HANDOVER DATE')
        shipby = _value_in(block, 'LATEST SHIP BY DATE')   # PDF 无连字符
        for st in _styles_in(block):
            for size, qty in zip(st['sizes'], st['qtys']):
                rows.append({
                    'ORDER NUMBER': order_no,
                    'DEVELOPMENT CODE': st['dev'],
                    'STYLE COLOUR CODE': st['scc'],
                    'STYLE DESCRIPTION': st['sdesc'],
                    'COLOUR DESCRIPTION': st['cdesc'],
                    'DESTINATION NAME': dest,
                    'SIZE': size,
                    'QUANTITY': qty,
                    'FINAL BOOKED BY DATE': booked,
                    'FINAL HANDOVER DATE': handover,
                    'LATEST SHIP-BY DATE': shipby,
                    'FREIGHT MODE': freight,
                })
    if not rows:
        raise ValueError('PDF 中未解析到任何订单明细行')
    return pd.DataFrame(rows, columns=PO_COLUMNS)


def load_po(path) -> pd.DataFrame:
    """按扩展名读取 PO：.pdf 走解析，其余按 CSV（utf-8 → gbk 回退）。"""
    if str(path).lower().endswith('.pdf'):
        return parse_po_pdf(path)
    try:
        return pd.read_csv(path, header=0)
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding='gbk', header=0)


# ── 下单价格表：客款号 → 厂款号 ────────────────────────────────────────
def load_fac_style_map(xlsx) -> tuple[dict, list]:
    """读「CO-下单价格表」，返回 ({客款号: 厂款号}, [冲突款号...])。

    客款号在表里常是数字型，两侧统一 str().strip()。同一客款号出现多次且
    厂款号不同 → 记入冲突列表（首个值生效），由调用方在汇总里提示。
    """
    if not xlsx:
        return {}, []
    df = pd.read_excel(xlsx, sheet_name=0)
    if '客款号' not in df.columns or '厂款号' not in df.columns:
        raise ValueError('价格表缺少「客款号」或「厂款号」列')

    out: dict[str, str] = {}
    conflicts: list[str] = []
    for style, fac in zip(df['客款号'], df['厂款号']):
        s = _clean_code(style)
        f = str(fac).strip() if fac is not None else ''
        if not s or not f or f.lower() == 'nan':
            continue
        if s in out:
            if out[s] != f and s not in conflicts:
                conflicts.append(s)
            continue
        out[s] = f
    return out, conflicts


def _clean_code(v) -> str:
    """数字型客款号 6340593.0 → '6340593'。"""
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


# ── 自检：PDF 解析结果必须与同名 CSV 完全一致 ──────────────────────────
if __name__ == '__main__':
    from pathlib import Path

    base = (Path(__file__).resolve().parent.parent / '测试数据' / '2026  Q2合同'
            / '3.19' / '6339064 - PO151435 - 002 Purchase Order')
    if not base.with_suffix('.pdf').exists():           # 目录名可能不同，兜底搜索
        root = Path(__file__).resolve().parent.parent / '测试数据'
        hits = list(root.rglob('6339064 - PO151435 - 002 Purchase Order.pdf'))
        base = hits[0].with_suffix('') if hits else base

    keys = ['ORDER NUMBER', 'DEVELOPMENT CODE', 'STYLE COLOUR CODE',
            'COLOUR DESCRIPTION', 'DESTINATION NAME', 'SIZE', 'QUANTITY',
            'FINAL BOOKED BY DATE', 'FINAL HANDOVER DATE',
            'LATEST SHIP-BY DATE', 'FREIGHT MODE']

    def _norm(df):
        return (df[keys].astype(str).sort_values(keys)
                .reset_index(drop=True))

    a = _norm(parse_po_pdf(base.with_suffix('.pdf')))
    b = _norm(pd.read_csv(base.with_suffix('.csv')))
    assert a.equals(b), a.compare(b).head(20)
    print('PDF 解析自检通过：%d 行与 CSV 完全一致 (%s)' % (len(a), base.name))
