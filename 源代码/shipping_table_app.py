#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""出货表导出工具 — 独立程序（第四功能，确认无误后再合并入主程序）。

输入：第三功能生成的「装箱单（PL）」xlsx 文件夹，或「点箱数明细」xlsx 文件（可同时选）。
流程：
  ① 选 PL 文件夹 / 点箱数明细文件 → 扫描，列出全部目的地和检测到的出货日期。
  ② 运行时分组：可把多个目的地「合并为一组」（如 BASH+南非）。
     同一组内，相同出货日期的 PO 自动合并为一个文件；
     不同日期则各生成一个文件（文件名含日期）。
  ③ 出厂日期和到深圳日期自动从 PL/点箱数明细数据读取；
     到深圳日期 = 出厂日期 +1 天（自动计算）。
  ④ 填编号日期/起始序号、输出文件夹 → 导出。

版式仿 测试数据/sample 1。生成逻辑见 shipping_table_export.py。
"""

import json
import sys
import traceback
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, simpledialog, ttk

import shipping_table_export as ste

if getattr(sys, 'frozen', False):
    APP_DIR = Path(sys.executable).parent
else:
    APP_DIR = Path(__file__).parent

CONFIG_FILE = APP_DIR / 'shipping_config.json'


def load_cfg() -> dict:
    try:
        return json.loads(CONFIG_FILE.read_text(encoding='utf-8'))
    except Exception:                                   # noqa: BLE001
        return {}


def save_cfg(**kw):
    cfg = load_cfg()
    cfg.update(kw)
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, ensure_ascii=False, indent=2),
                               encoding='utf-8')
    except Exception:                                   # noqa: BLE001
        pass


def build_into(parent):
    cfg = load_cfg()
    p = {'padx': 8, 'pady': 4}
    parent.columnconfigure(1, weight=1)

    state = {'by_dest': None}

    # ── 文件夹行 ──────────────────────────────────────────────────
    # Row 0: PL folder
    tk.Label(parent, text='PL（装箱单）文件夹：').grid(row=0, column=0, sticky='e', **p)
    fvar_pl = tk.StringVar(value=cfg.get('pl_folder', ''))
    tk.Entry(parent, textvariable=fvar_pl).grid(row=0, column=1, sticky='ew', **p)
    def _browse_pl():
        path = filedialog.askdirectory()
        if path:
            fvar_pl.set(path); save_cfg(pl_folder=path)
    tk.Button(parent, text='文件夹…', command=_browse_pl).grid(row=0, column=2, **p)

    # Row 1: Output folder
    tk.Label(parent, text='输出文件夹：').grid(row=1, column=0, sticky='e', **p)
    fvar_out = tk.StringVar(value=cfg.get('out_dir', ''))
    tk.Entry(parent, textvariable=fvar_out).grid(row=1, column=1, sticky='ew', **p)
    def _browse_out():
        path = filedialog.askdirectory()
        if path:
            fvar_out.set(path); save_cfg(out_dir=path)
    tk.Button(parent, text='文件夹…', command=_browse_out).grid(row=1, column=2, **p)

    # ── 编号参数 ────────────────────────────────────────────────
    pf = tk.Frame(parent)
    pf.grid(row=2, column=0, columnspan=3, sticky='w', **p)
    tk.Label(pf, text='编号日期(YYYYMMDD)：').grid(row=0, column=0, sticky='e',
                                                    padx=4, pady=3)
    code_date_var = tk.StringVar(
        value=cfg.get('code_date') or date.today().strftime('%Y%m%d'))
    tk.Entry(pf, textvariable=code_date_var, width=10).grid(
        row=0, column=1, sticky='w', padx=4)
    tk.Label(pf, text='起始序号：').grid(row=0, column=2, sticky='e', padx=(16, 4))
    code_start_var = tk.StringVar(value=cfg.get('code_start', '1'))
    tk.Entry(pf, textvariable=code_start_var, width=4).grid(
        row=0, column=3, sticky='w', padx=4)
    tk.Label(pf, text='  （出厂/到深圳日期、星期自动从数据读取）',
             fg='gray').grid(row=0, column=4, sticky='w')

    # ── 扫描按钮 ────────────────────────────────────────────────
    def do_scan():
        folder = fvar_pl.get().strip()
        if not folder or not Path(folder).is_dir():
            messagebox.showwarning('提示', '请先选择有效的 PL 文件夹')
            return
        save_cfg(pl_folder=folder)
        try:
            by_dest = ste.parse_pl_folder(folder)
        except Exception as e:                          # noqa: BLE001
            messagebox.showerror('扫描失败', f'{e}\n\n{traceback.format_exc()}')
            return
        if not by_dest:
            messagebox.showwarning('提示', '未在该文件夹解析到任何 PL 数据')
            return
        state['by_dest'] = by_dest
        reset_groups_one_per_dest()
        refresh_dest_list()

    tk.Button(parent, text='① 扫描 PL', width=14, command=do_scan).grid(
        row=3, column=0, sticky='w', **p)
    tk.Label(parent,
             text='  扫描后在下方把目的地拼成分组；同组内每个出货日期自动生成一个文件',
             fg='gray').grid(row=3, column=1, columnspan=2, sticky='w')

    # ── 双栏布局：目的地列表 | 分组 ─────────────────────────────
    body = tk.Frame(parent)
    body.grid(row=4, column=0, columnspan=3, sticky='nsew', padx=8, pady=4)
    parent.rowconfigure(4, weight=1)
    body.columnconfigure(0, weight=1)
    body.columnconfigure(2, weight=2)
    body.rowconfigure(1, weight=1)

    tk.Label(body, text='检测到的目的地（可多选）').grid(row=0, column=0, sticky='w')
    dest_lb = tk.Listbox(body, selectmode='multiple', height=14,
                         exportselection=False, font=('宋体', 10))
    dest_lb.grid(row=1, column=0, sticky='nsew')

    dest_click_order: list[int] = []

    def _on_dest_click(event):
        idx = dest_lb.nearest(event.y)
        if idx < 0 or idx >= dest_lb.size():
            return 'break'
        if dest_lb.selection_includes(idx):
            dest_lb.selection_clear(idx)
            if idx in dest_click_order:
                dest_click_order.remove(idx)
        else:
            dest_lb.selection_set(idx)
            if idx not in dest_click_order:
                dest_click_order.append(idx)
        return 'break'

    dest_lb.bind('<Button-1>', _on_dest_click)

    mid = tk.Frame(body)
    mid.grid(row=1, column=1, padx=8)
    tk.Button(mid, text='合并所选 →', width=12,
              command=lambda: merge_selected()).pack(pady=6)
    tk.Button(mid, text='每个单独', width=12,
              command=lambda: reset_groups_one_per_dest()).pack(pady=6)
    tk.Button(mid, text='← 删除分组', width=12,
              command=lambda: remove_group()).pack(pady=6)
    tk.Button(mid, text='改标签', width=12,
              command=lambda: rename_group()).pack(pady=6)

    tk.Label(body,
             text='分组（每组×每日期 = 一个出货表文件；双击改标签）').grid(
        row=0, column=2, sticky='w')
    grp_cols = ('label', 'dests', 'dates', 'pos', 'boxes')
    grp_tv = ttk.Treeview(body, columns=grp_cols, show='headings', height=14)
    for c, h, w in zip(grp_cols,
                        ['表头标签', '目的地', '日期', 'PO', '箱'],
                        [220, 200, 100, 45, 45]):
        grp_tv.heading(c, text=h)
        grp_tv.column(c, width=w,
                      anchor='w' if c in ('label', 'dests', 'dates') else 'center')
    grp_tv.grid(row=1, column=2, sticky='nsew')
    grp_tv.bind('<Double-1>', lambda e: rename_group())

    groups: list[dict] = []   # list of {'label': str, 'dests': [str]}

    def _dest_keys():
        return list((state['by_dest'] or {}).keys())

    def refresh_dest_list():
        dest_click_order.clear()
        dest_lb.delete(0, 'end')
        for dest, lst in (state['by_dest'] or {}).items():
            dates = sorted({s.ship_date for s in lst if s.ship_date},
                           key=ste._date_sort_key)
            boxes = sum(s.total_boxes for s in lst)
            date_str = ' / '.join(dates) if dates else '(未填日期)'
            dest_lb.insert('end', f'{dest}   [{date_str}]   ({len(lst)}PO/{boxes}箱)')

    def refresh_group_tv():
        grp_tv.delete(*grp_tv.get_children())
        by = state['by_dest'] or {}
        for i, g in enumerate(groups):
            ships = [s for d in g['dests'] for s in by.get(d, [])]
            dates = sorted({s.ship_date for s in ships if s.ship_date},
                           key=ste._date_sort_key)
            date_str = ' / '.join(dates) if dates else '(未填)'
            pos = len(ships)
            boxes = sum(s.total_boxes for s in ships)
            grp_tv.insert('', 'end', iid=str(i),
                          values=(g['label'], ' + '.join(g['dests']),
                                  date_str, pos, boxes))

    def reset_groups_one_per_dest():
        groups.clear()
        for dest in (state['by_dest'] or {}):
            groups.append({'label': dest, 'dests': [dest]})
        refresh_group_tv()

    def merge_selected():
        if len(dest_click_order) < 2:
            messagebox.showinfo('提示', '请按顺序逐个点选至少 2 个目的地再合并')
            return
        keys = _dest_keys()
        dests = [keys[i] for i in dest_click_order if 0 <= i < len(keys)]
        dest_click_order.clear()
        dest_lb.selection_clear(0, 'end')
        for g in groups:
            g['dests'] = [d for d in g['dests'] if d not in dests]
        groups[:] = [g for g in groups if g['dests']]
        groups.append({'label': ' + '.join(dests), 'dests': dests})
        refresh_group_tv()

    def _selected_idx():
        sel = grp_tv.selection()
        return int(sel[0]) if sel else None

    def remove_group():
        i = _selected_idx()
        if i is None:
            messagebox.showinfo('提示', '请先在右侧选中一个分组')
            return
        g = groups.pop(i)
        for d in g['dests']:
            groups.append({'label': d, 'dests': [d]})
        refresh_group_tv()

    def rename_group():
        i = _selected_idx()
        if i is None:
            messagebox.showinfo('提示', '请先在右侧选中一个分组')
            return
        new = simpledialog.askstring('表头标签', '出货表表头(H1)文本：',
                                     initialvalue=groups[i]['label'])
        if new:
            groups[i]['label'] = new.strip()
            refresh_group_tv()

    # ── 导出 ────────────────────────────────────────────────────
    def do_export():
        if not state['by_dest']:
            messagebox.showwarning('提示', '请先扫描数据')
            return
        if not groups:
            messagebox.showwarning('提示', '没有可导出的分组')
            return
        out_dir = fvar_out.get().strip()
        if not out_dir:
            messagebox.showwarning('提示', '请选择输出文件夹')
            return
        code_date = code_date_var.get().strip() or date.today().strftime('%Y%m%d')
        try:
            code_start = int(code_start_var.get().strip() or '1')
        except ValueError:
            code_start = 1
        specs = [ste.GroupSpec(label=g['label'], destinations=list(g['dests']))
                 for g in groups]
        save_cfg(out_dir=out_dir, code_date=code_date, code_start=str(code_start))
        try:
            written = ste.export_shipping_tables(
                state['by_dest'], out_dir, specs, code_date, code_start)
            written_dk = ste.export_dianka_workbooks(
                state['by_dest'], out_dir, specs)
        except Exception as e:                          # noqa: BLE001
            messagebox.showerror('导出失败', f'{e}\n\n{traceback.format_exc()}')
            return
        messagebox.showinfo(
            '完成',
            f'已导出到：\n{out_dir}\n\n'
            f'出货表 {len(written)} 个：\n' + '\n'.join(written)
            + f'\n\n点箱数明细 {len(written_dk)} 个：\n' + '\n'.join(written_dk))

    btnrow = tk.Frame(parent)
    btnrow.grid(row=6, column=0, columnspan=3, sticky='w', **p)
    tk.Button(btnrow, text='② 导出出货表 + 点箱数明细', width=24, command=do_export,
              bg='#2d7d46', fg='white').pack(side='left')
    tk.Button(btnrow, text='打开输出文件夹',
              command=lambda: _open_dir(fvar_out.get())).pack(
        side='left', padx=10)


def _open_dir(path):
    try:
        if path and Path(path).is_dir():
            import os
            os.startfile(path)                          # noqa: S606
    except Exception:                                   # noqa: BLE001
        pass


def main():
    root = tk.Tk()
    root.title('出货表导出工具（测试版）')
    root.geometry('1020x720')
    root.minsize(900, 600)
    build_into(root)
    root.mainloop()


if __name__ == '__main__':
    main()
