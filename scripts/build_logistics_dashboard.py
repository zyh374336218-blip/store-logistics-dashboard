"""
Build logistics dashboard data from datasource/daily.xlsx (preferred) or daily.csv.

Usage:
  python scripts/build_logistics_dashboard.py
  python scripts/build_logistics_dashboard.py --source datasource/daily.xlsx

Fails fast if any required field is empty or PH NO is duplicated.

Column order (xlsx):
  market, PID, PH NO, DATE, ORDER-SHIPPED(days), Uploaded Weight,
  Actual Weighed Weight, qty, Weight Variance
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import re
import sys
from collections import Counter
from datetime import date, datetime, timezone
from io import StringIO
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_SOURCE = ROOT / "datasource" / "daily.xlsx"
OUT_JS = ROOT / "dashboard" / "data" / "logistics.js"
OUT_JSON = ROOT / "dashboard" / "data" / "logistics.json"

WEIGHT_OUTLIER_THRESHOLD = 0.2


class DataSourceError(Exception):
    """Raised when the source file fails quality checks."""


def parse_date(value) -> str:
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")
    s = str(value).strip()
    for fmt in ("%Y.%m.%d", "%Y-%m-%d", "%Y/%m/%d", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    m = re.fullmatch(r"(\d{4})[./-](\d{1,2})[./-](\d{1,2})", s)
    if m:
        y, mo, d = map(int, m.groups())
        return f"{y:04d}-{mo:02d}-{d:02d}"
    raise DataSourceError(f"无法解析 DATE: {value!r}")


def cell_str(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def require_str(value, field: str, row_no: int) -> str:
    s = cell_str(value)
    if s == "" or s.upper() in {"#N/A", "N/A", "NA", "NULL", "NONE", "-"}:
        raise DataSourceError(f"第 {row_no} 行字段「{field}」为空，数据源有误")
    return s


def to_float(value, field: str, row_no: int) -> float:
    if value is None or (isinstance(value, str) and value.strip() == ""):
        raise DataSourceError(f"第 {row_no} 行字段「{field}」为空，数据源有误")
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        v = float(value)
    else:
        s = str(value).strip()
        if s.upper() in {"#N/A", "N/A", "NA", "NULL", "NONE", "-"}:
            raise DataSourceError(f"第 {row_no} 行字段「{field}」为空，数据源有误")
        try:
            v = float(s)
        except ValueError as e:
            raise DataSourceError(
                f"第 {row_no} 行字段「{field}」不是数值: {value!r}"
            ) from e
    if not math.isfinite(v):
        raise DataSourceError(f"第 {row_no} 行字段「{field}」无效: {value!r}")
    return v


def read_xlsx_rows(path: Path) -> list[list]:
    try:
        import openpyxl
    except ImportError as e:
        raise DataSourceError("需要 openpyxl，请先 pip install openpyxl") from e

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    # drop fully empty trailing rows
    cleaned = []
    for row in rows:
        if row is None:
            continue
        if all(c is None or str(c).strip() == "" for c in row):
            continue
        cleaned.append(row)
    return cleaned


def read_csv_rows(path: Path) -> list[list[str]]:
    raw = path.read_bytes()
    text = None
    for enc in ("utf-8-sig", "utf-8", "gbk", "cp1252", "latin1"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise DataSourceError(f"无法解码文件编码: {path}")
    lines = [ln for ln in text.splitlines() if ln.strip()]
    if not lines:
        raise DataSourceError("数据源为空")
    return list(csv.reader(StringIO("\n".join(lines))))


def read_source_rows(path: Path) -> list[list]:
    suffix = path.suffix.lower()
    if suffix in {".xlsx", ".xlsm"}:
        return read_xlsx_rows(path)
    if suffix == ".csv":
        return read_csv_rows(path)
    raise DataSourceError(f"不支持的数据源类型: {path.suffix}")


def load_orders(path: Path) -> tuple[list[dict], dict]:
    rows = read_source_rows(path)
    if len(rows) < 2:
        raise DataSourceError("数据源无明细行")
    # Expect 9 columns with qty before Weight Variance
    if len(rows[0]) < 9:
        raise DataSourceError(
            f"表头列数不足 9（需含 qty 与 Weight Variance），实际 {len(rows[0])}"
        )

    orders: list[dict] = []
    ph_nos: list[str] = []

    for i, row in enumerate(rows[1:], start=2):
        if len(row) < 9:
            raise DataSourceError(f"第 {i} 行列数不足 9，数据源有误")
        market = require_str(row[0], "market", i)
        pid = require_str(row[1], "PID", i)
        ph_no = require_str(row[2], "PH NO", i)
        if isinstance(row[3], (date, datetime)):
            date_s = parse_date(row[3])
        else:
            date_s = parse_date(require_str(row[3], "DATE", i))
        shipped = to_float(row[4], "ORDER-SHIPPED（days）", i)
        uploaded = to_float(row[5], "Uploaded Weight", i)
        actual = to_float(row[6], "Actual Weighed Weight", i)
        qty = to_float(row[7], "qty", i)
        # Weight Variance from source file — used as-is for View E / outlier KPIs
        variance = to_float(row[8], "Weight Variance", i)

        is_outlier = variance > WEIGHT_OUTLIER_THRESHOLD
        orders.append(
            {
                "market": market,
                "pid": pid,
                "ph_no": ph_no,
                "date": date_s,
                "order_shipped_days": round(shipped, 2),
                "uploaded_weight_kg": round(uploaded, 2),
                "actual_weight_kg": round(actual, 2),
                "qty": round(qty, 2),
                "weight_variance_kg": round(variance, 2),
                "is_weight_outlier": is_outlier,
            }
        )
        ph_nos.append(ph_no)

    dup = sorted(k for k, v in Counter(ph_nos).items() if v > 1)
    if dup:
        preview = ", ".join(dup[:10])
        more = f" 等共 {len(dup)} 个" if len(dup) > 10 else ""
        raise DataSourceError(
            f"存在重复 PH NO，数据源有误: {preview}{more}"
        )

    dates = sorted({o["date"] for o in orders})
    try:
        rel = str(path.relative_to(ROOT)).replace("\\", "/")
    except ValueError:
        rel = str(path)
    meta = {
        "source": rel,
        "period_month": dates[0][:7] if dates else None,
        "period_start": dates[0] if dates else None,
        "period_end": dates[-1] if dates else None,
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "order_count": len(orders),
        "weight_outlier_threshold": WEIGHT_OUTLIER_THRESHOLD,
        "note": "Weight Variance is taken from source as-is; qty is units per PH NO",
    }
    return orders, meta


def write_outputs(orders: list[dict], meta: dict) -> None:
    OUT_JS.parent.mkdir(parents=True, exist_ok=True)
    payload = {"meta": meta, "orders": orders}
    OUT_JSON.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    js = (
        "/* Auto-generated by scripts/build_logistics_dashboard.py — do not edit */\n"
        f"window.LOGISTICS_DATA = {json.dumps(payload, ensure_ascii=False)};\n"
    )
    OUT_JS.write_text(js, encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Build logistics dashboard data")
    parser.add_argument(
        "--source",
        type=Path,
        default=DEFAULT_SOURCE,
        help="Path to daily.xlsx or daily.csv",
    )
    # backward-compatible alias
    parser.add_argument("--csv", type=Path, default=None, help=argparse.SUPPRESS)
    args = parser.parse_args()
    src = args.csv or args.source
    path = src if src.is_absolute() else ROOT / src
    if not path.exists():
        print(f"[error] 找不到数据源: {path}", file=sys.stderr)
        return 1

    try:
        orders, meta = load_orders(path)
        write_outputs(orders, meta)
    except DataSourceError as e:
        print(f"[error] 数据源有误: {e}", file=sys.stderr)
        return 1

    print(
        f"[ok] {meta['order_count']} orders | "
        f"{meta['period_start']} ~ {meta['period_end']} | "
        f"wrote {OUT_JS.relative_to(ROOT).as_posix()} "
        f"and {OUT_JSON.relative_to(ROOT).as_posix()}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
