# -*- coding: utf-8 -*-
"""验货报告「交接」清单(manifest)读写 —— 跟单工具(导出端) 与 QC 工具(导入端) 共用。

⚠️  本文件在两个程序里保持完全一致(逐字复制),避免协议漂移。

约定:
  共享盘根目录下,每次导出生成一个「批次文件夹」,内含:
    - 若干 .xlsm 验货报告(待验)
    - BOM物料分类汇总.xlsx
    - manifest.json   (本文件定义;导出端写一次,导入端只读)
    - 已完成\          (QC 完成后把成品另存到此;凭文件是否存在判断完成状态,
                        从而避免双方同时改写 manifest 造成的网络盘写冲突)

manifest.json 结构:
  {
    "batch_id": str, "created": iso, "bom_file": str|None,
    "reports": [
      {"file","style","mode"(inline|final),"label","supplier","factory",
       "total_order_qty","produced_qty","needle_file"(可选,同款共用一份断针记录)}, ...
    ]
  }
"""
import json
import os
from datetime import datetime

MANIFEST_NAME = "manifest.json"
DONE_DIRNAME = "已完成"
BOM_FILENAME = "BOM物料分类汇总.xlsx"


def build_manifest(batch_id, reports, bom_file=None):
    return {
        "batch_id": batch_id,
        "created": datetime.now().isoformat(timespec="seconds"),
        "bom_file": bom_file,
        "reports": list(reports),
    }


def write_manifest(batch_dir, manifest):
    """原子写入 manifest.json(先写临时文件再 replace)。"""
    os.makedirs(batch_dir, exist_ok=True)
    path = os.path.join(batch_dir, MANIFEST_NAME)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)
    os.replace(tmp, path)
    return path


def read_manifest(batch_dir):
    path = os.path.join(batch_dir, MANIFEST_NAME)
    if not os.path.isfile(path):
        return None
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def done_dir(batch_dir):
    return os.path.join(batch_dir, DONE_DIRNAME)


def done_files(batch_dir):
    """已完成\ 内的文件名集合(用于判断完成状态)。"""
    d = done_dir(batch_dir)
    if not os.path.isdir(d):
        return set()
    try:
        return {n for n in os.listdir(d) if not n.startswith("~$")}
    except OSError:
        return set()


def is_done(report_file, done_set):
    return report_file in done_set


def find_batch_dirs(root):
    """在交接根目录下查找所有「批次文件夹」(含 manifest.json)。
    支持两层结构:
      - 深度 1:root/<批次>            (旧的扁平结构,向后兼容)
      - 深度 2:root/<年份>/<批次>     (现行:按年份分文件夹)
    跳过「已完成」目录;有界查找(不深层递归),在网络盘上更快、更稳。
    返回批次目录的绝对路径列表。
    """
    found = []
    if not os.path.isdir(root):
        return found
    try:
        children = os.listdir(root)
    except OSError:
        return found
    for name in children:
        if name == DONE_DIRNAME:
            continue
        d1 = os.path.join(root, name)
        if not os.path.isdir(d1):
            continue
        if os.path.isfile(os.path.join(d1, MANIFEST_NAME)):
            found.append(d1)            # 深度 1:扁平批次
            continue
        try:
            sub = os.listdir(d1)         # 深度 2:年份/批次
        except OSError:
            continue
        for sname in sub:
            if sname == DONE_DIRNAME:
                continue
            d2 = os.path.join(d1, sname)
            if os.path.isdir(d2) and os.path.isfile(os.path.join(d2, MANIFEST_NAME)):
                found.append(d2)
    return found
