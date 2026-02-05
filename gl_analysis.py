#!/usr/bin/env python3
"""Basic GL journal entry analysis without third-party dependencies."""

from __future__ import annotations

import argparse
import csv
import math
import os
import re
import statistics
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from datetime import datetime, timedelta
from typing import Dict, List, Optional, Tuple

EXCEL_EPOCH = datetime(1899, 12, 30)
DATE_COLUMN_HINTS = {"date"}
NUMERIC_COLUMN_HINTS = {"amount", "debit", "credit", "balance"}
PREFERRED_CATEGORY_COLUMNS = [
    "Source",
    "BusinessUnit",
    "PreparerID",
    "AccountType",
    "AccountClass",
]
PREFERRED_NUMERIC_COLUMNS = ["Amount", "Debit", "Credit", "AbsoluteAmount"]


def col_letters(cell_ref: str) -> str:
    match = re.match(r"([A-Z]+)", cell_ref)
    return match.group(1) if match else ""


def col_to_index(letters: str) -> int:
    idx = 0
    for ch in letters:
        idx = idx * 26 + (ord(ch) - 64)
    return idx - 1


def read_shared_strings(zf: zipfile.ZipFile) -> List[str]:
    if "xl/sharedStrings.xml" not in zf.namelist():
        return []
    root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    ns = {"ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    strings: List[str] = []
    for si in root.findall("ns:si", ns):
        texts = []
        for t in si.findall(".//ns:t", ns):
            texts.append(t.text or "")
        strings.append("".join(texts))
    return strings


def read_sheet(zf: zipfile.ZipFile, sheet_path: str) -> List[List[Optional[str]]]:
    shared_strings = read_shared_strings(zf)
    root = ET.fromstring(zf.read(sheet_path))
    ns = {"ns": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    rows: List[Dict[str, Optional[str]]] = []
    for row in root.findall(".//ns:sheetData/ns:row", ns):
        row_vals: Dict[str, Optional[str]] = {}
        for cell in row.findall("ns:c", ns):
            cell_ref = cell.attrib.get("r")
            if not cell_ref:
                continue
            cell_type = cell.attrib.get("t")
            value_node = cell.find("ns:v", ns)
            value = value_node.text if value_node is not None else None
            if cell_type == "s" and value is not None:
                value = shared_strings[int(value)]
            elif cell_type == "inlineStr":
                inline_node = cell.find("ns:is/ns:t", ns)
                value = inline_node.text if inline_node is not None else ""
            row_vals[cell_ref] = value
        rows.append(row_vals)

    max_col = 0
    for row_vals in rows:
        for cell_ref in row_vals:
            col = col_to_index(col_letters(cell_ref))
            if col > max_col:
                max_col = col

    num_cols = max_col + 1
    table: List[List[Optional[str]]] = []
    for row_vals in rows:
        row = [None] * num_cols
        for cell_ref, value in row_vals.items():
            col = col_to_index(col_letters(cell_ref))
            row[col] = value
        table.append(row)

    return table


def load_workbook(path: str) -> Tuple[List[str], List[List[Optional[str]]]]:
    with zipfile.ZipFile(path) as zf:
        sheet_path = "xl/worksheets/sheet1.xml"
        table = read_sheet(zf, sheet_path)
    headers = [h if h is not None else "" for h in table[0]]
    data = table[1:]
    return headers, data


def is_numeric(value: Optional[str]) -> bool:
    if value is None:
        return False
    try:
        float(value)
        return True
    except (TypeError, ValueError):
        return False


def to_float(value: Optional[str]) -> Optional[float]:
    if value is None:
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def excel_serial_to_date(serial: float) -> Optional[datetime]:
    if serial < 0 or serial > 60000:
        return None
    try:
        return EXCEL_EPOCH + timedelta(days=serial)
    except OverflowError:
        return None


def infer_column_type(name: str, values: List[Optional[str]]) -> str:
    lowered = name.lower()
    if any(hint in lowered for hint in DATE_COLUMN_HINTS):
        return "date"
    numeric_values = [v for v in values if is_numeric(v)]
    if numeric_values:
        if any(hint in lowered for hint in NUMERIC_COLUMN_HINTS):
            return "numeric"
        if len(numeric_values) / max(len(values), 1) > 0.6:
            return "numeric"
    return "text"


def summarize_numeric(values: List[float]) -> Dict[str, float]:
    if not values:
        return {}
    return {
        "min": min(values),
        "max": max(values),
        "mean": statistics.mean(values),
        "median": statistics.median(values),
    }


def summarize_text(values: List[str]) -> Dict[str, int]:
    counter = Counter(values)
    return {"unique": len(counter)}


def summarize_date(values: List[datetime]) -> Dict[str, datetime]:
    if not values:
        return {}
    return {"min": min(values), "max": max(values)}


def format_number(value: Optional[float]) -> str:
    if value is None:
        return "n/a"
    return f"{value:,.2f}"


def write_svg_bar_chart(path: str, title: str, labels: List[str], values: List[int]) -> None:
    width = 900
    height = 500
    margin = 80
    max_value = max(values) if values else 1
    bar_width = (width - 2 * margin) / max(len(values), 1)

    def bar_height(val: int) -> float:
        return (height - 2 * margin) * (val / max_value)

    svg_parts = [
        f"<svg xmlns='http://www.w3.org/2000/svg' width='{width}' height='{height}'>",
        f"<rect width='100%' height='100%' fill='white' />",
        f"<text x='{width/2}' y='{margin/2}' font-size='20' text-anchor='middle'>{title}</text>",
    ]

    for idx, (label, val) in enumerate(zip(labels, values)):
        x = margin + idx * bar_width
        bh = bar_height(val)
        y = height - margin - bh
        svg_parts.append(
            f"<rect x='{x}' y='{y}' width='{bar_width*0.8}' height='{bh}' fill='#4C78A8' />"
        )
        svg_parts.append(
            f"<text x='{x + bar_width*0.4}' y='{height - margin + 20}' font-size='10' text-anchor='middle' transform='rotate(20 {x + bar_width*0.4},{height - margin + 20})'>{label}</text>"
        )
        svg_parts.append(
            f"<text x='{x + bar_width*0.4}' y='{y - 5}' font-size='10' text-anchor='middle'>{val}</text>"
        )

    svg_parts.append("</svg>")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(svg_parts))


def write_svg_histogram(path: str, title: str, bins: List[Tuple[float, float, int]]) -> None:
    width = 900
    height = 500
    margin = 80
    max_value = max((count for _, _, count in bins), default=1)
    bar_width = (width - 2 * margin) / max(len(bins), 1)

    svg_parts = [
        f"<svg xmlns='http://www.w3.org/2000/svg' width='{width}' height='{height}'>",
        f"<rect width='100%' height='100%' fill='white' />",
        f"<text x='{width/2}' y='{margin/2}' font-size='20' text-anchor='middle'>{title}</text>",
    ]

    for idx, (start, end, count) in enumerate(bins):
        x = margin + idx * bar_width
        bh = (height - 2 * margin) * (count / max_value)
        y = height - margin - bh
        svg_parts.append(
            f"<rect x='{x}' y='{y}' width='{bar_width*0.8}' height='{bh}' fill='#F58518' />"
        )
        label = f"{start:,.0f}-{end:,.0f}"
        svg_parts.append(
            f"<text x='{x + bar_width*0.4}' y='{height - margin + 20}' font-size='9' text-anchor='middle' transform='rotate(20 {x + bar_width*0.4},{height - margin + 20})'>{label}</text>"
        )
        svg_parts.append(
            f"<text x='{x + bar_width*0.4}' y='{y - 5}' font-size='10' text-anchor='middle'>{count}</text>"
        )

    svg_parts.append("</svg>")
    with open(path, "w", encoding="utf-8") as f:
        f.write("\n".join(svg_parts))


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Analyze GL journal XLSX exports.")
    parser.add_argument(
        "--input",
        default="je_samples (1).xlsx",
        help="Path to the XLSX file to analyze.",
    )
    parser.add_argument(
        "--output",
        default="analysis_output",
        help="Output directory for summary and charts.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    source_path = args.input
    if not os.path.exists(source_path):
        raise FileNotFoundError(f"Missing {source_path}")

    headers, data = load_workbook(source_path)
    row_count = len(data)
    col_count = len(headers)

    output_dir = args.output
    os.makedirs(output_dir, exist_ok=True)

    columns: Dict[str, List[Optional[str]]] = {name: [] for name in headers}
    for row in data:
        for idx, name in enumerate(headers):
            columns[name].append(row[idx] if idx < len(row) else None)

    summary_lines = []
    summary_lines.append("Journal Entry Data Summary")
    summary_lines.append("==========================")
    summary_lines.append(f"Source file: {source_path}")
    summary_lines.append(f"Rows: {row_count}")
    summary_lines.append(f"Columns: {col_count}")
    summary_lines.append("")

    date_columns: Dict[str, List[datetime]] = {}
    numeric_columns: Dict[str, List[float]] = {}
    text_columns: Dict[str, List[str]] = {}

    for name, values in columns.items():
        column_type = infer_column_type(name, values)
        non_null = [v for v in values if v not in (None, "")]
        summary_lines.append(f"Column: {name}")
        summary_lines.append(f"  Type: {column_type}")
        summary_lines.append(f"  Non-null: {len(non_null)}")
        summary_lines.append(f"  Nulls: {len(values) - len(non_null)}")

        if column_type == "date":
            parsed_dates = []
            for v in non_null:
                if is_numeric(v):
                    parsed = excel_serial_to_date(float(v))
                    if parsed:
                        parsed_dates.append(parsed)
                else:
                    try:
                        parsed_dates.append(datetime.fromisoformat(str(v)))
                    except ValueError:
                        continue
            if parsed_dates:
                summary = summarize_date(parsed_dates)
                summary_lines.append(
                    f"  Date range: {summary['min'].date()} to {summary['max'].date()}"
                )
            date_columns[name] = parsed_dates
        elif column_type == "numeric":
            parsed_numbers = [to_float(v) for v in non_null]
            parsed_numbers = [v for v in parsed_numbers if v is not None]
            summary = summarize_numeric(parsed_numbers)
            summary_lines.append(
                "  Stats: "
                f"min={format_number(summary.get('min'))}, "
                f"max={format_number(summary.get('max'))}, "
                f"mean={format_number(summary.get('mean'))}, "
                f"median={format_number(summary.get('median'))}"
            )
            numeric_columns[name] = parsed_numbers
        else:
            cleaned = [str(v).strip() for v in non_null if str(v).strip()]
            summary = summarize_text(cleaned)
            summary_lines.append(f"  Unique values: {summary.get('unique', 0)}")
            text_columns[name] = cleaned

        summary_lines.append("")

    summary_lines.append("Top categorical fields")
    summary_lines.append("-----------------------")
    for name, values in list(text_columns.items())[:5]:
        counter = Counter(values)
        most_common = counter.most_common(5)
        summary_lines.append(f"{name}:")
        for label, count in most_common:
            summary_lines.append(f"  {label}: {count}")
        summary_lines.append("")

    summary_path = os.path.join(output_dir, "summary.txt")
    with open(summary_path, "w", encoding="utf-8") as f:
        f.write("\n".join(summary_lines))

    # Create a bar chart for a categorical column with manageable cardinality.
    chart_written = False
    for preferred in PREFERRED_CATEGORY_COLUMNS:
        if preferred in text_columns:
            values = text_columns[preferred]
            counts = Counter(values)
            if 2 <= len(counts) <= 50:
                labels, vals = zip(*counts.most_common(10))
                chart_path = os.path.join(output_dir, "top_categories.svg")
                write_svg_bar_chart(
                    chart_path, f"Top {preferred} Values", list(labels), list(vals)
                )
                chart_written = True
                break

    if not chart_written:
        for name, values in text_columns.items():
            counts = Counter(values)
            if 2 <= len(counts) <= 20:
                labels, vals = zip(*counts.most_common(10))
                chart_path = os.path.join(output_dir, "top_categories.svg")
                write_svg_bar_chart(
                    chart_path, f"Top {name} Values", list(labels), list(vals)
                )
                chart_written = True
                break

    if not chart_written and text_columns:
        name, values = next(iter(text_columns.items()))
        counts = Counter(values)
        labels, vals = zip(*counts.most_common(10))
        chart_path = os.path.join(output_dir, "top_categories.svg")
        write_svg_bar_chart(chart_path, f"Top {name} Values", list(labels), list(vals))

    # Create histogram for a numeric column.
    hist_written = False
    for preferred in PREFERRED_NUMERIC_COLUMNS:
        if preferred in numeric_columns and len(numeric_columns[preferred]) > 10:
            name = preferred
            values = numeric_columns[preferred]
            min_val = min(values)
            max_val = max(values)
            if max_val == min_val:
                continue
            bin_count = 10
            bin_size = (max_val - min_val) / bin_count
            bins = []
            for i in range(bin_count):
                start = min_val + i * bin_size
                end = start + bin_size
                bins.append([start, end, 0])
            for val in values:
                idx = min(int((val - min_val) / bin_size), bin_count - 1)
                bins[idx][2] += 1
            hist_path = os.path.join(output_dir, "numeric_histogram.svg")
            write_svg_histogram(hist_path, f"Distribution of {name}", bins)
            hist_written = True
            break

    if not hist_written:
        for name, values in numeric_columns.items():
            if len(values) > 10:
                min_val = min(values)
                max_val = max(values)
                if max_val == min_val:
                    continue
                bin_count = 10
                bin_size = (max_val - min_val) / bin_count
                bins = []
                for i in range(bin_count):
                    start = min_val + i * bin_size
                    end = start + bin_size
                    bins.append([start, end, 0])
                for val in values:
                    idx = min(int((val - min_val) / bin_size), bin_count - 1)
                    bins[idx][2] += 1
                hist_path = os.path.join(output_dir, "numeric_histogram.svg")
                write_svg_histogram(hist_path, f"Distribution of {name}", bins)
                hist_written = True
                break

    if not hist_written and numeric_columns:
        name, values = next(iter(numeric_columns.items()))
        if values:
            min_val = min(values)
            max_val = max(values)
            bin_count = 10
            bin_size = (max_val - min_val) / bin_count if max_val != min_val else 1
            bins = []
            for i in range(bin_count):
                start = min_val + i * bin_size
                end = start + bin_size
                bins.append([start, end, 0])
            for val in values:
                idx = min(int((val - min_val) / bin_size), bin_count - 1)
                bins[idx][2] += 1
            hist_path = os.path.join(output_dir, "numeric_histogram.svg")
            write_svg_histogram(hist_path, f"Distribution of {name}", bins)

    # Write a CSV with basic column metadata for re-use.
    metadata_path = os.path.join(output_dir, "column_metadata.csv")
    with open(metadata_path, "w", encoding="utf-8", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["column", "inferred_type", "non_null", "nulls"])
        for name, values in columns.items():
            column_type = infer_column_type(name, values)
            non_null = [v for v in values if v not in (None, "")]
            writer.writerow([name, column_type, len(non_null), len(values) - len(non_null)])


if __name__ == "__main__":
    main()
