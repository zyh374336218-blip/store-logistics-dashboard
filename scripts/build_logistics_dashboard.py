"""
Build logistics dashboard data from datasource/daily.csv.

Usage:
  python scripts/build_logistics_dashboard.py
  python scripts/build_logistics_dashboard.py --csv datasource/daily.csv

Fails fast if any required field is empty or PH NO is duplicated.
"""

from __future__ import annotations

import argparse
import json
import math
import re
import sys
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_CSV = ROOT / "datasource" / "daily.csv"
OUT_JS = ROOT / "dashboard" / "data" / "logistics.js"
OUT_JSON = ROOT / "dashboard" / "data" / "logistics.json"

STANDARD_COLS = [
    "market",
    "pid",
    "ph_no",
    "date",
    "order_shipped_days",
    "uploaded_weight_kg",
    "actual_weight_kg",
    "weight_variance_kg",
]

WEIGHT_OUTLIER_THRESHOLD = 0.2
VARIANCE_TOLERANCE = 0.05


class DataSourceError(Exception):
    """Raised when the source CSV fails quality checks."""


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

    rows: list[list[str]] = []
    for line in lines:
        rows.append(next(csv_reader_line(line)))
    return rows


def csv_reader_line(line: str):
    import csv
    from io import StringIO

    yield from csv.reader(StringIO(line))


def parse_date(value: str) -> str:
    s = value.strip()
    for fmt in ("%Y.%m.%d", "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    # 2026.7.1 style
    m = re.fullmatch(r"(\d{4})[./-](\d{1,2})[./-](\d{1,2})", s)
    if m:
        y, mo, d = map(int, m.groups())
        return f"{y:04d}-{mo:02d}-{d:02d}"
    raise DataSourceError(f"无法解析 DATE: {value!r}")


def to_float(value: str, field: str, row_no: int) -> float:
    s = (value or "").strip()
    if s == "":
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


def require_str(value: str, field: str, row_no: int) -> str:
    s = (value or "").strip()
    if s == "":
        raise DataSourceError(f"第 {row_no} 行字段「{field}」为空，数据源有误")
    return s


def load_orders(path: Path) -> tuple[list[dict], dict]:
    rows = read_csv_rows(path)
    if len(rows) < 2:
        raise DataSourceError("数据源无明细行")
    if len(rows[0]) < 8:
        raise DataSourceError(f"表头列数不足 8，实际 {len(rows[0])}")

    orders: list[dict] = []
    ph_nos: list[str] = []
    variance_mismatches: list[str] = []

    for i, row in enumerate(rows[1:], start=2):
        if len(row) < 8:
            raise DataSourceError(f"第 {i} 行列数不足 8，数据源有误")
        # skip fully blank trailing lines already filtered; still check empties
        market = require_str(row[0], "market", i)
        pid = require_str(row[1], "PID", i)
        ph_no = require_str(row[2], "PH NO", i)
        date = parse_date(require_str(row[3], "DATE", i))
        shipped = to_float(row[4], "ORDER-SHIPPED（days）", i)
        uploaded = to_float(row[5], "Uploaded Weight", i)
        actual = to_float(row[6], "Actual Weighed Weight", i)
        variance = to_float(row[7], "Weight Variance", i)

        expected = actual - uploaded
        if abs(expected - variance) > VARIANCE_TOLERANCE:
            variance_mismatches.append(
                f"第 {i} 行 PH NO={ph_no}: Variance={variance} "
                f"≠ Actual-Uploaded={expected:.4f}"
            )

        is_outlier = variance > WEIGHT_OUTLIER_THRESHOLD
        orders.append(
            {
                "market": market,
                "pid": pid,
                "ph_no": ph_no,
                "date": date,
                "order_shipped_days": round(shipped, 2),
                "uploaded_weight_kg": round(uploaded, 2),
                "actual_weight_kg": round(actual, 2),
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

    if variance_mismatches:
        # Soft warning to stderr; do not block build (source may round differently)
        print(
            f"[warn] Weight Variance 与 Actual-Uploaded 不一致 "
            f"{len(variance_mismatches)} 行（容差 {VARIANCE_TOLERANCE}）",
            file=sys.stderr,
        )
        for line in variance_mismatches[:5]:
            print(f"  {line}", file=sys.stderr)

    dates = sorted({o["date"] for o in orders})
    meta = {
        "source": str(path.relative_to(ROOT)).replace("\\", "/"),
        "period_month": dates[0][:7] if dates else None,
        "period_start": dates[0] if dates else None,
        "period_end": dates[-1] if dates else None,
        "generated_at": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "order_count": len(orders),
        "weight_outlier_threshold": WEIGHT_OUTLIER_THRESHOLD,
        "note": "演示样例可被正式全量 CSV 替换后重跑本脚本",
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
        "--csv",
        type=Path,
        default=DEFAULT_CSV,
        help="Path to daily.csv",
    )
    args = parser.parse_args()
    csv_path = args.csv if args.csv.is_absolute() else ROOT / args.csv
    if not csv_path.exists():
        print(f"[error] 找不到数据源: {csv_path}", file=sys.stderr)
        return 1

    try:
        orders, meta = load_orders(csv_path)
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
