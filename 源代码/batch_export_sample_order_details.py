"""Batch-run the CO PO import tool with representative field-config sample data."""

from __future__ import annotations

import argparse
import importlib.util
import re
import sys
from pathlib import Path

import pandas as pd


COLOUR_CN = {
    "BLACK": "黑色",
    "BLACK/BLACK/WHITE/WHITE": "黑色/黑色/白色/白色",
    "BLACK/CREAM": "黑色/奶油色",
    "CHEEKY CHERRY": "樱桃红",
    "COCONUT MILK": "椰奶白",
    "COCONUT MILK/PINK KISS/MILKY MATCHA": "椰奶白/亲吻粉/抹茶绿",
    "CREAM": "奶油色",
    "MERRY BERRY SPOT": "莓果波点",
    "MILKY MATCHA": "抹茶绿",
    "PINK KISS": "亲吻粉",
    "PINK KISS/TONAL ELASTIC": "亲吻粉/同色松紧带",
    "SOFT MOCHA": "柔和摩卡",
    "SOFT MOCHA / TONAL ELASTIC": "柔和摩卡/同色松紧带",
    "STEELY BLUE WASH/TONAL ELASTIC": "钢蓝水洗/同色松紧带",
    "VIOLET BLUE SPOT": "紫蓝波点",
}


def load_tool(module_path: Path):
    spec = importlib.util.spec_from_file_location("po_import_app", module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Cannot import tool module: {module_path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def read_csv(path: Path) -> pd.DataFrame:
    try:
        return pd.read_csv(path, header=0)
    except UnicodeDecodeError:
        return pd.read_csv(path, encoding="gbk", header=0)


def clean(value) -> str:
    text = str(value).strip()
    return "" if text in {"", "nan", "None", "NaT"} else text


def po_number(df: pd.DataFrame, path: Path) -> str:
    if "ORDER NUMBER" in df.columns and not df.empty:
        value = clean(df["ORDER NUMBER"].iloc[0])
        if value:
            return value
    match = re.search(r"(?:^| - )(PO[^ -]+)(?= - )", path.name.replace("\u00a0", " "))
    return match.group(1) if match else path.stem


def first_value(df: pd.DataFrame, column: str) -> str:
    if column not in df.columns:
        return ""
    for value in df[column]:
        value = clean(value)
        if value:
            return value
    return ""


def build_port_map(port_path: Path) -> dict[str, tuple[str, str]]:
    records = pd.read_excel(port_path, header=0).to_dict("records")
    result = {}
    for entry in records:
        source = clean(entry.get("对应CO合同港口"))
        if not source:
            continue
        result[source.split(" - ")[0].strip()] = (
            clean(entry.get("名称")),
            clean(entry.get("英文名称")),
        )
    return result


def destination_map(df: pd.DataFrame, ports: dict[str, tuple[str, str]]):
    result = {}
    if "DESTINATION NAME" not in df.columns:
        return result
    for destination in dict.fromkeys(df["DESTINATION NAME"].dropna().astype(str)):
        code = destination.split(" - ")[0].strip()
        result[destination] = ports.get(code, ("", ""))
    return result


def factory_date(tool, df: pd.DataFrame, style_code: str) -> str:
    if "FINAL HANDOVER DATE" not in df.columns:
        return ""
    subset = df
    if "STYLE COLOUR CODE" in df.columns:
        subset = df[df["STYLE COLOUR CODE"].astype(str) == style_code]
    parsed = [tool.App._parse_date(value) for value in subset["FINAL HANDOVER DATE"]]
    parsed = sorted(value for value in parsed if value)
    return parsed[0] if parsed else ""


def colour_map(tool, df: pd.DataFrame, sample_ids: dict[str, str], fac_style: str):
    result = {}
    if "STYLE COLOUR CODE" not in df.columns:
        return result
    descriptions = df.get("COLOUR DESCRIPTION", pd.Series([""] * len(df)))
    for style_code, description in zip(df["STYLE COLOUR CODE"], descriptions):
        style_code = clean(style_code)
        description = clean(description)
        if style_code and style_code not in result:
            result[style_code] = (
                sample_ids.setdefault(description, f"SMP-C{len(sample_ids) + 1:03d}"),
                COLOUR_CN.get(description, f"样例颜色{len(sample_ids):03d}"),
                description,
                fac_style,
                factory_date(tool, df, style_code),
            )
    return result


def export_one(
    tool,
    csv_path: Path,
    template_path: Path,
    output_dir: Path,
    ports: dict[str, tuple[str, str]],
    sample_ids: dict[str, str],
) -> list[Path]:
    df = read_csv(csv_path)
    number = po_number(df, csv_path)
    development_code = first_value(df, "DEVELOPMENT CODE") or number
    fac_style = f"样例-{development_code}"

    app = tool.App.__new__(tool.App)
    app.df = df
    app.sizes = (
        sorted(df["SIZE"].astype(str).str.strip().unique().tolist(), key=tool._size_sort_key)
        if "SIZE" in df.columns
        else []
    )
    configured_colours = colour_map(tool, df, sample_ids, fac_style)
    configured_destinations = destination_map(df, ports)
    app._colour_map = lambda: configured_colours
    app._dest_map = lambda: configured_destinations

    groups = app._file_groups()
    split_mode = len(groups) > 1
    saved = []
    for index, group in enumerate(groups):
        rows = app._build_rows(style_codes=set(group) if split_mode else None)
        if not rows:
            continue
        for row_number, row in enumerate(rows, start=1):
            row["流水号"] = row_number
        suffix = "" if index == 0 else f"_{index + 1}"
        output_path = output_dir / f"{number}_{fac_style}{suffix}.xlsx"
        app._write_excel(str(template_path), rows, str(output_path))
        saved.append(output_path)
    return saved


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--source", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--tool-module", type=Path, required=True)
    parser.add_argument("--template", type=Path, required=True)
    parser.add_argument("--ports", type=Path, required=True)
    args = parser.parse_args()

    if args.output.exists() and any(args.output.iterdir()):
        raise RuntimeError(f"Output folder is not empty: {args.output}")
    args.output.mkdir(parents=True, exist_ok=True)

    tool = load_tool(args.tool_module)
    # These fields are part of the configured template's post-size section.
    # The current GUI tool omits them from KNOWN_FIXED_COLS, which causes
    # dynamic-size adjustment to remove 比例/数量合计 from generated files.
    tool.KNOWN_FIXED_COLS.update({"工厂出厂期", "厂款号"})
    ports = build_port_map(args.ports)
    sample_ids: dict[str, str] = {}
    csv_files = sorted(args.source.glob("*Purchase Order.csv"))

    exported = []
    for index, csv_path in enumerate(csv_files, start=1):
        files = export_one(
            tool, csv_path, args.template, args.output, ports, sample_ids
        )
        exported.extend(files)
        print(f"[{index:02d}/{len(csv_files):02d}] {csv_path.name} -> {len(files)} file(s)")

    print(f"PO CSV files: {len(csv_files)}")
    print(f"Exported workbooks: {len(exported)}")
    print(f"Output folder: {args.output}")


if __name__ == "__main__":
    main()
