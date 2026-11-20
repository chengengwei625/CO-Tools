"""
stamp_cartons_tab.py
--------------------
build_into(parent) adapter — embeds SSCC carton-stamping UI as a notebook tab.
Engine: stamp_cartons.py (PyMuPDF-based).
"""

import json
import sys
import threading
from collections import defaultdict
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox

from stamp_cartons import (
    load_csv_cartons,
    read_all_label_pages,
    validate_labels_vs_cartons,
    assign_carton_numbers,
    stamp_pdfs,
    export_xlsx,
)

if getattr(sys, "frozen", False):
    _APP_DIR = Path(sys.executable).parent
else:
    _APP_DIR = Path(__file__).parent

_CONFIG_FILE = _APP_DIR / "stamp_config.json"

ACCENT  = "#1A73E8"
TEXT    = "#202124"
MUTED   = "#5F6368"
SUCCESS = "#188038"
WARNING = "#E37400"
ERROR   = "#C5221F"


def _load_cfg() -> dict:
    try:
        return json.loads(_CONFIG_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_cfg(src: str, lbl: str, out: str, pl: str) -> None:
    try:
        _CONFIG_FILE.write_text(
            json.dumps({"src": src, "lbl": lbl, "out": out, "pl": pl}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception:
        pass


def build_into(parent: tk.Widget) -> None:
    cfg = _load_cfg()

    parent.columnconfigure(0, weight=1)
    parent.rowconfigure(1, weight=1)

    # ── Controls ──────────────────────────────────────────────────
    ctrl = tk.Frame(parent)
    ctrl.grid(row=0, column=0, sticky="ew", padx=14, pady=10)
    ctrl.columnconfigure(1, weight=1)

    src_var = tk.StringVar(value=cfg.get("src", ""))
    pl_var  = tk.StringVar(value=cfg.get("pl", ""))
    lbl_var = tk.StringVar(value=cfg.get("lbl", ""))
    out_var = tk.StringVar(value=cfg.get("out", ""))

    fields = [
        ("CSV 文件夹（箱单明细）：",    src_var),
        ("装箱单文件夹（xlsx）：",      pl_var),
        ("标签文件夹（PDF）：",         lbl_var),
        ("输出文件夹：",                out_var),
    ]

    def _browse(var: tk.StringVar):
        d = filedialog.askdirectory(initialdir=var.get() or ".")
        if d:
            var.set(d)

    for row, (label, var) in enumerate(fields):
        tk.Label(ctrl, text=label, anchor="e").grid(
            row=row, column=0, sticky="e", padx=(0, 8), pady=4)
        tk.Entry(ctrl, textvariable=var).grid(
            row=row, column=1, sticky="ew", ipady=3, pady=4)
        tk.Button(ctrl, text="浏览…",
                  command=lambda v=var: _browse(v)).grid(
            row=row, column=2, padx=(6, 0), pady=4)

    run_btn = tk.Button(
        ctrl, text="▶  开始标注",
        font=("Segoe UI", 11, "bold"),
        fg="white", bg=ACCENT,
        activebackground="#1557B0", activeforeground="white",
        relief="flat", padx=20, pady=7, cursor="hand2",
    )
    run_btn.grid(row=len(fields), column=0, columnspan=3, pady=(10, 4))

    progress = ttk.Progressbar(ctrl, mode="indeterminate", length=500)
    progress.grid(row=len(fields) + 1, column=0, columnspan=3, sticky="ew", pady=(0, 2))
    progress.grid_remove()

    # ── Log ───────────────────────────────────────────────────────
    log_box = scrolledtext.ScrolledText(
        parent, font=("Consolas", 9), state="disabled",
        bg="#FAFAFA", fg=TEXT, relief="solid", bd=1, wrap="word",
    )
    log_box.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 10))
    log_box.tag_config("warn",    foreground=WARNING)
    log_box.tag_config("error",   foreground=ERROR)
    log_box.tag_config("success", foreground=SUCCESS)
    log_box.tag_config("head",    foreground=ACCENT, font=("Consolas", 9, "bold"))
    log_box.tag_config("muted",   foreground=MUTED)

    running = [False]

    def _log(msg: str, tag: str = ""):
        log_box.configure(state="normal")
        log_box.insert("end", msg + "\n", tag or ())
        log_box.see("end")
        log_box.configure(state="disabled")

    def _ask(title, message) -> bool:
        """从工作线程弹确认框（编组回主线程并等待结果）。"""
        holder = {}
        ev = threading.Event()

        def _do():
            holder["ok"] = messagebox.askyesno(title, message)
            ev.set()
        parent.after(0, _do)
        ev.wait()
        return holder.get("ok", False)

    def _worker(src: Path, pl: Path, lbl: Path, out: Path):
        def log(msg, tag=""):
            parent.after(0, _log, msg, tag)

        try:
            log("正在加载 CSV 箱单数据…", "head")
            csv_cartons = load_csv_cartons(src)
            for po, boxes in csv_cartons.items():
                log(f"  {po:20s}  {len(boxes):3d} 箱", "muted")

            log("\n正在读取标签 PDF 页面…", "head")
            records = read_all_label_pages(
                lbl,
                log_fn=lambda m: log(m, "warn" if "WARNING" in m else "muted"),
            )
            log(
                f"  共 {len(records)} 页，跨 "
                f"{len(set(r['file'] for r in records))} 个 PDF 文件",
                "muted",
            )

            # ── 处理前预校验：装箱单(CSV) 与标签数量是否一致 ──
            log("\n正在校验装箱单与标签数量…", "head")
            chk = validate_labels_vs_cartons(records, csv_cartons)
            if not (chk["missing_csv"] or chk["count_mismatch"] or chk["total_mismatch"]):
                log("  ✓ 各 PO 的装箱单箱数与标签数量一致", "success")
            for po in chk["missing_csv"]:
                log(f"  ⚠ PO {po} 在标签中出现，但没有对应的装箱单(CSV)，将无法标注箱号", "warn")
            for po, npages, nboxes in chk["count_mismatch"]:
                log(f"  ⚠ PO {po}：标签 {npages} 页 ≠ 装箱单 {nboxes} 箱", "warn")
            for fname, n, npages in chk["total_mismatch"]:
                log(f"  ⚠ {fname}：标签标示 of {n} ≠ 实际 {npages} 页", "warn")

            if chk["missing_csv"]:
                pos = "、".join(chk["missing_csv"])
                if not _ask(
                    "装箱单缺失",
                    f"以下 PO 在标签中出现，但数据源中没有对应的装箱单(CSV)：\n\n{pos}\n\n"
                    "这些 PO 的箱号将无法标注（显示为 ?）。是否仍要继续？",
                ):
                    log("\n已取消：请补齐缺失的装箱单(CSV)后再处理。", "error")
                    return

            log("\n正在分配箱号…", "head")
            records, warnings = assign_carton_numbers(
                records, csv_cartons,
                log_fn=lambda m: log(m, "warn"),
            )

            log("\n正在标注并保存 PDF…", "head")
            stamp_pdfs(records, out, log_fn=lambda m: log(m, "muted"))

            log("\n正在导出 SSCC Excel 清单…", "head")
            export_xlsx(
                records, csv_cartons, out,
                log_fn=lambda m: log(m, "muted"),
                source_dir=pl if pl.is_dir() else src,
            )

            stamped = sum(1 for r in records if r.get("ctn_number", "?") != "?")
            per_po: dict[str, int] = defaultdict(int)
            for r in records:
                if r.get("ctn_number", "?") != "?":
                    per_po[r["po"]] += 1

            log("\n=== 完成 ===", "head")
            log(f"  已标注页数：{stamped} / {len(records)}", "success")
            log(f"  警告数量：{len(warnings)}", "warn" if warnings else "muted")
            log(f"  输出文件夹：{out}", "muted")
            log("\n  各订单明细：", "muted")
            for po, count in per_po.items():
                log(f"    {po:20s}  {count} 箱", "muted")

        except Exception as exc:
            log(f"\n错误：{exc}", "error")

        finally:
            running[0] = False
            parent.after(0, progress.stop)
            parent.after(0, progress.grid_remove)
            parent.after(0, run_btn.configure, {"state": "normal"})

    def _run():
        if running[0]:
            return
        running[0] = True
        run_btn.configure(state="disabled")
        log_box.configure(state="normal")
        log_box.delete("1.0", "end")
        log_box.configure(state="disabled")
        progress.grid()
        progress.start(12)

        src = Path(src_var.get())
        pl  = Path(pl_var.get())
        lbl = Path(lbl_var.get())
        out = Path(out_var.get())
        _save_cfg(str(src), str(lbl), str(out), str(pl))

        threading.Thread(target=_worker, args=(src, pl, lbl, out), daemon=True).start()

    run_btn.configure(command=_run)
