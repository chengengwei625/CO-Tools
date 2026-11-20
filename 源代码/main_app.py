"""
集杰跟单工具 — 合并启动器
将所有功能整合在同一窗口的标签页中。
"""
import tkinter as tk
from tkinter import ttk
from app_version import APP_TITLE

import po_import_app       # ① 订单明细导入
import generate_tabs       # ② 港口/数量汇总
import order_summary_app   # ③ 订单汇总
import qc_report_app       # ④ 验货报告 › 验货报告
import needle_report_app   # ④ 验货报告 › 断针记录
import bom_summary_app     # ④ 验货报告 › BOM分类
import qc_status_app       # ④ 验货报告 › 验货进度
import packing_list_app    # ⑤ 装箱单生成
import shipping_table_app  # ⑥ 出货表导出
import stamp_cartons_tab   # ⑦ 装箱贴纸（SSCC 标注）
import asos_labels_app     # ASOS ASN 标签及箱单


def main():
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry('1200x900')
    root.minsize(960, 720)

    nb = ttk.Notebook(root)
    nb.pack(fill='both', expand=True)

    # ── Tab 1: 订单明细导入 ───────────────────────────────────────
    tab1 = ttk.Frame(nb)
    nb.add(tab1, text='  ① 订单明细导入  ')
    po_import_app.App(tab1)

    # ── Tab 2: 港口/数量汇总 ───────────────────────────────────────
    tab2 = ttk.Frame(nb)
    nb.add(tab2, text='  ② 港口/数量汇总  ')
    generate_tabs.build_into(tab2)

    # ── Tab 3: 订单汇总 ───────────────────────────────────────────
    tab3 = ttk.Frame(nb)
    nb.add(tab3, text='  ③ 订单汇总  ')
    order_summary_app.build_into(tab3)

    # ── Tab 4: 验货报告（含断针记录 + BOM分类 子标签）──────────────
    tab4 = ttk.Frame(nb)
    nb.add(tab4, text='  ④ 验货报告  ')

    # Top bar: "generate all" sits above the sub-notebook
    top_bar = tk.Frame(tab4)
    top_bar.pack(side='top', fill='x', padx=8, pady=(6, 2))
    all_status = tk.StringVar(value='')
    tk.Label(top_bar, textvariable=all_status, fg='#005599').pack(side='right', padx=8)

    sub_nb = ttk.Notebook(tab4)
    sub_nb.pack(fill='both', expand=True)

    sub_qc = ttk.Frame(sub_nb)
    sub_nb.add(sub_qc, text='  验货报告  ')
    do_qc = qc_report_app.build_into(sub_qc)

    sub_needle = ttk.Frame(sub_nb)
    sub_nb.add(sub_needle, text='  断针记录  ')
    do_needle = needle_report_app.build_into(sub_needle)

    sub_bom = ttk.Frame(sub_nb)
    sub_nb.add(sub_bom, text='  BOM分类  ')
    do_bom = bom_summary_app.build_into(sub_bom)

    sub_status = ttk.Frame(sub_nb)
    sub_nb.add(sub_status, text='  验货进度  ')
    qc_status_app.build_into(sub_status)

    def _generate_all():
        all_status.set('正在生成全部报告…')
        if do_qc:
            do_qc()
        if do_needle:
            do_needle()
        if do_bom:
            do_bom()

    tk.Button(top_bar, text='  生成全部报告  ', command=_generate_all,
              bg='#1a5276', fg='white', font=('Arial', 11, 'bold')).pack(side='left')

    # ── Tab 5: 装箱单生成 ─────────────────────────────────────────
    tab5 = ttk.Frame(nb)
    nb.add(tab5, text='  ⑤ 装箱单生成  ')
    packing_list_app.build_into(tab5)

    # ── Tab 6: 出货表导出 + 点箱数明细 ───────────────────────────
    tab6 = ttk.Frame(nb)
    nb.add(tab6, text='  ⑥ 出货表导出  ')
    shipping_table_app.build_into(tab6)

    # ── Tab 7: 装箱贴纸（SSCC 标签箱号标注）────────────────────
    tab7 = ttk.Frame(nb)
    nb.add(tab7, text='  ⑦ 装箱贴纸  ')
    stamp_cartons_tab.build_into(tab7)

    tab_asos = ttk.Frame(nb)
    nb.add(tab_asos, text='  ASOS 标签/箱单  ')
    asos_labels_app.build_into(tab_asos)

    root.mainloop()


if __name__ == '__main__':
    main()
