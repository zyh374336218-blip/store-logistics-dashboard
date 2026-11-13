"""
Build logistics dashboard data from datasource/daily.xlsx (preferred) or daily.csv.

Usage:
  python scripts/build_logistics_dashboard.py
  python scripts/build_logistics_dashboard.py --source datasource/daily.xlsx

Fails fast if any required field is empty or PH NO is duplicated.

Required columns (by header name; order flexible):
  market, PID, PH NO, DATE, ORDER-SHIPPED(days), Uploaded Weight,
  Actual Weighed Weight, Weight Variance
Optional: qty (defaults to 1 when absent)
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
UNDELIVERED_OVERDUE_DAYS = 5
NOT_DELIVERED_SHEET_CANDIDATES = ("not delievered", "not delivered", "Not Delievered")


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


def _clean_rows(rows: list) -> list[list]:
    cleaned = []
    for row in rows:
        if row is None:
            continue
        if all(c is None or str(c).strip() == "" for c in row):
            continue
        cleaned.append(list(row))
    return cleaned


def read_xlsx_sheet_rows(path: Path, sheet_name: str | None = None) -> list[list]:
    try:
        import openpyxl
    except ImportError as e:
        raise DataSourceError("需要 openpyxl，请先 pip install openpyxl") from e

    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        if sheet_name is None:
            ws = wb[wb.sheetnames[0]]
        elif sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
        else:
            raise DataSourceError(
                f"找不到工作表「{sheet_name}」，现有: {wb.sheetnames}"
            )
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    return _clean_rows(rows)


def read_xlsx_rows(path: Path) -> list[list]:
    return read_xlsx_sheet_rows(path, sheet_name=None)


def find_not_delivered_sheet(path: Path) -> str | None:
    try:
        import openpyxl
    except ImportError as e:
        raise DataSourceError("需要 openpyxl，请先 pip install openpyxl") from e
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        names = wb.sheetnames
    finally:
        wb.close()
    lower_map = {n.lower(): n for n in names}
    for cand in NOT_DELIVERED_SHEET_CANDIDATES:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    for n in names:
        key = n.lower().replace(" ", "")
        if "notdelievered" in key or "notdelivered" in key:
            return n
    return None


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
        # Prefer sheet named "daily" when present
        try:
            import openpyxl

            wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
            names = wb.sheetnames
            wb.close()
            daily_name = None
            for n in names:
                if n.lower() == "daily":
                    daily_name = n
                    break
            return read_xlsx_sheet_rows(path, sheet_name=daily_name)
        except DataSourceError:
            raise
        except Exception:
            return read_xlsx_rows(path)
    if suffix == ".csv":
        return read_csv_rows(path)
    raise DataSourceError(f"不支持的数据源类型: {path.suffix}")


def load_not_delivered(path: Path) -> list[dict]:
    """Load undelivered orders from the not-delivered sheet (xlsx only)."""
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return []
    sheet = find_not_delivered_sheet(path)
    if sheet is None:
        print(
            "[warn] 未找到 not delievered / not delivered 工作表，跳过未送达模块",
            file=sys.stderr,
        )
        return []
    rows = read_xlsx_sheet_rows(path, sheet_name=sheet)
    if len(rows) < 2:
        return []
    # Expected: ph_nr, id_partner, ..., order_pdate, shipped_date, ..., today, order_pdate-today
    header = [cell_str(h).lower() for h in rows[0]]
    def col(*names: str) -> int | None:
        for i, h in enumerate(header):
            for name in names:
                if h == name or h.replace(" ", "") == name.replace(" ", ""):
                    return i
        return None

    i_ph = col("ph_nr", "ph no", "ph_no")
    i_pid = col("id_partner", "pid")
    i_pdate = col("order_pdate")
    i_today = col("today")
    i_days = col("order_pdate-today", "order_pdate_today")
    if i_ph is None or i_pid is None or i_days is None:
        raise DataSourceError(
            f"工作表「{sheet}」缺少必要列 ph_nr / id_partner / order_pdate-today，实际表头: {rows[0]}"
        )

    def _is_blank_or_formula(v) -> bool:
        if v is None:
            return True
        if isinstance(v, float) and not math.isfinite(v):
            return True
        if isinstance(v, str):
            s = v.strip()
            return s == "" or s.startswith("=") or s.lower() in {"null", "#n/a", "n/a"}
        return False

    out: list[dict] = []
    fallback_today = datetime.now(timezone.utc).date()
    for i, row in enumerate(rows[1:], start=2):
        if max(i_ph, i_pid, i_days) >= len(row):
            raise DataSourceError(f"未送达表第 {i} 行列数不足，数据源有误")
        ph_no = require_str(row[i_ph], "ph_nr", i)
        pid = require_str(row[i_pid], "id_partner", i)
        order_pdate = None
        today_s = None
        if i_pdate is not None and i_pdate < len(row) and row[i_pdate] not in (None, ""):
            try:
                order_pdate = parse_date(row[i_pdate])
            except DataSourceError:
                order_pdate = cell_str(row[i_pdate]) or None
        if i_today is not None and i_today < len(row) and not _is_blank_or_formula(
            row[i_today]
        ):
            try:
                today_s = parse_date(row[i_today])
            except DataSourceError:
                today_s = cell_str(row[i_today]) or None

        days_raw = row[i_days] if i_days < len(row) else None
        if _is_blank_or_formula(days_raw):
            if not order_pdate:
                raise DataSourceError(
                    f"未送达表第 {i} 行 order_pdate-today 为空且无法用 order_pdate 推算"
                )
            pdate = date.fromisoformat(order_pdate)
            tdate = (
                date.fromisoformat(today_s) if today_s else fallback_today
            )
            if today_s is None:
                today_s = tdate.isoformat()
            days = float((tdate - pdate).days)
        else:
            days = to_float(days_raw, "order_pdate-today", i)
            if today_s is None:
                today_s = fallback_today.isoformat()

        out.append(
            {
                "ph_no": ph_no,
                "pid": pid,
                "order_pdate": order_pdate,
                "today": today_s,
                "order_pdate_today_days": round(days, 2),
                "is_overdue_gt5": days > UNDELIVERED_OVERDUE_DAYS,
            }
        )
    return out


def _norm_header(value) -> str:
    s = cell_str(value).lower()
    s = s.replace("（", "(").replace("）", ")")
    s = re.sub(r"\s+", " ", s).strip()
    return s


def _header_index_map(header_row: list) -> dict[str, int]:
    aliases = {
        "market": ("market",),
        "pid": ("pid",),
        "ph_no": ("ph no", "ph_no", "phno"),
        "date": ("date",),
        "order_shipped_days": (
            "order-shipped(days)",
            "order-shipped (days)",
            "order shipped (days)",
            "order-shipped（days)",
        ),
        "uploaded_weight": ("uploaded weight",),
        "actual_weight": ("actual weighed weight", "actual weight"),
        "qty": ("qty", "quantity"),
        "weight_variance": ("weight variance",),
    }
    by_norm = {_norm_header(h): i for i, h in enumerate(header_row) if cell_str(h)}
    out: dict[str, int] = {}
    for key, names in aliases.items():
        for name in names:
            if name in by_norm:
                out[key] = by_norm[name]
                break
    return out


def _cell(row: list, idx: int | None):
    if idx is None or idx >= len(row):
        return None
    return row[idx]


def load_orders(path: Path) -> tuple[list[dict], dict]:
    rows = read_source_rows(path)
    if len(rows) < 2:
        raise DataSourceError("数据源无明细行")

    col = _header_index_map(rows[0])
    required = [
        "market",
        "pid",
        "ph_no",
        "date",
        "order_shipped_days",
        "uploaded_weight",
        "actual_weight",
        "weight_variance",
    ]
    missing = [k for k in required if k not in col]
    if missing:
        raise DataSourceError(
            f"表头缺少必要列: {', '.join(missing)}；实际表头: {rows[0]}"
        )

    orders: list[dict] = []
    ph_nos: list[str] = []
    qty_idx = col.get("qty")

    for i, row in enumerate(rows[1:], start=2):
        market = require_str(_cell(row, col["market"]), "market", i)
        pid = require_str(_cell(row, col["pid"]), "PID", i)
        ph_no = require_str(_cell(row, col["ph_no"]), "PH NO", i)
        date_raw = _cell(row, col["date"])
        if isinstance(date_raw, (date, datetime)):
            date_s = parse_date(date_raw)
        else:
            date_s = parse_date(require_str(date_raw, "DATE", i))
        shipped = to_float(
            _cell(row, col["order_shipped_days"]), "ORDER-SHIPPED（days）", i
        )
        uploaded = to_float(
            _cell(row, col["uploaded_weight"]), "Uploaded Weight", i
        )
        actual = to_float(
            _cell(row, col["actual_weight"]), "Actual Weighed Weight", i
        )
        # Prefer source Weight Variance; if blank/formula cache missing, use Actual−Uploaded.
        variance_raw = _cell(row, col["weight_variance"])
        if (
            variance_raw is None
            or (isinstance(variance_raw, float) and not math.isfinite(variance_raw))
            or (isinstance(variance_raw, str) and variance_raw.strip() == "")
            or (isinstance(variance_raw, str) and variance_raw.strip().startswith("="))
        ):
            variance = actual - uploaded
        else:
            variance = to_float(variance_raw, "Weight Variance", i)
        if qty_idx is None:
            qty = 1.0
        else:
            qty = to_float(_cell(row, qty_idx), "qty", i)

        # Outlier uses raw precision; UI display still formats to 2 decimals.
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
                "weight_variance_kg": variance,
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
        "note": "Weight Variance is taken from source as-is; qty defaults to 1 when absent",
    }
    return orders, meta


def write_outputs(
    orders: list[dict],
    meta: dict,
    not_delivered: list[dict] | None = None,
) -> None:
    OUT_JS.parent.mkdir(parents=True, exist_ok=True)
    nd = not_delivered or []
    overdue = sum(1 for x in nd if x.get("is_overdue_gt5"))
    meta = {
        **meta,
        "not_delivered_count": len(nd),
        "not_delivered_overdue_gt5_count": overdue,
        "not_delivered_overdue_threshold_days": UNDELIVERED_OVERDUE_DAYS,
    }
    payload = {"meta": meta, "orders": orders, "not_delivered": nd}
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
        not_delivered = load_not_delivered(path)
        write_outputs(orders, meta, not_delivered)
    except DataSourceError as e:
        print(f"[error] 数据源有误: {e}", file=sys.stderr)
        return 1

    print(
        f"[ok] {meta['order_count']} orders | "
        f"{meta['period_start']} ~ {meta['period_end']} | "
        f"not_delivered={len(not_delivered)} "
        f"overdue_gt5={sum(1 for x in not_delivered if x.get('is_overdue_gt5'))} | "
        f"wrote {OUT_JS.relative_to(ROOT).as_posix()} "
        f"and {OUT_JSON.relative_to(ROOT).as_posix()}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
