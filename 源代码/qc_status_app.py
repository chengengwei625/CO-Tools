# -*- coding: utf-8 -*-
"""验货进度 —— 读取交接共享盘各批次的 manifest + 「已完成」目录,显示完成/待验。

跟单端凭此看到 QC 端处理进度。完成状态由「已完成\」内是否存在同名文件判断,
不改写 manifest,避免与导出端/多 QC 端的写冲突。
"""
import os
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, ttk

import generate_tabs as gt
import qc_handoff as hf


def build_into(parent):
    cfg = gt.load_saved()
    p = {'padx': 8, 'pady': 4}
    parent.columnconfigure(1, weight=1)
    parent.rowconfigure(3, weight=1)

    tk.Label(parent, text='交接根目录：').grid(row=0, column=0, sticky='e', **p)
    root_var = tk.StringVar(value=cfg.get('qc_out', '') or '')
    tk.Entry(parent, textvariable=root_var).grid(row=0, column=1, sticky='ew', **p)

    def _pick():
        d = filedialog.askdirectory()
        if d:
            root_var.set(d)
    tk.Button(parent, text='文件夹…', command=_pick).grid(row=0, column=2, **p)

    summary = tk.StringVar(value='')
    tk.Label(parent, textvariable=summary, fg='#005599').grid(
        row=1, column=0, columnspan=3, sticky='w', **p)

    cols = ('batch', 'style', 'type', 'status')
    tree = ttk.Treeview(parent, columns=cols, show='headings')
    for c, t, w in (('batch', '批次', 200), ('style', '客款号', 130),
                    ('type', '类型', 140), ('status', '状态', 110)):
        tree.heading(c, text=t)
        tree.column(c, width=w, anchor='w')
    tree.grid(row=3, column=0, columnspan=3, sticky='nsew', **p)
    tree.tag_configure('done', foreground='#2d7d46')
    tree.tag_configure('pending', foreground='#b00000')

    def refresh():
        for i in tree.get_children():
            tree.delete(i)
        root = root_var.get().strip()
        if not root or not Path(root).is_dir():
            summary.set('请选择有效的交接根目录')
            return
        batches = total = done = 0
        bdirs = hf.find_batch_dirs(root)   # 支持 年份/批次 两层结构
        bdirs.sort(key=lambda d: os.path.relpath(d, root), reverse=True)
        for bdir in bdirs:
            m = hf.read_manifest(bdir)
            if not m:
                continue
            batches += 1
            blabel = os.path.relpath(bdir, root)   # 如 2026\2026-06-30_143000
            dset = hf.done_files(bdir)
            for r in m.get('reports', []):
                is_done = r.get('file') in dset
                total += 1
                done += 1 if is_done else 0
                label = r.get('label') or ('FI' if r.get('mode') == 'final' else 'IL')
                typ = ('尾期 ' if r.get('mode') == 'final' else '中期 ') + label
                tree.insert('', 'end',
                            values=(blabel, r.get('style', ''), typ,
                                    '✓ 已完成' if is_done else '待验'),
                            tags=('done' if is_done else 'pending',))
        summary.set('共 %d 批次 · %d 份报告 · 已完成 %d · 待验 %d'
                    % (batches, total, done, total - done))

    btns = tk.Frame(parent)
    btns.grid(row=2, column=0, columnspan=3, sticky='w', **p)
    tk.Button(btns, text='  🔄 刷新进度  ', command=refresh,
              bg='#1a5276', fg='white', font=('Arial', 10, 'bold')).pack(side='left')

    if root_var.get().strip():
        parent.after(200, refresh)
    return refresh


def main():
    root = tk.Tk()
    root.title('验货进度')
    root.geometry('760x520')
    build_into(root)
    root.mainloop()


if __name__ == '__main__':
    main()
