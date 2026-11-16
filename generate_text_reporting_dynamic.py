# -*- coding: utf-8 -*-
"""FBB Area-4 PAMASUKA - Text Reporting FULL REV FINAL 6.0

Sumber:
- Control Data: angka pivot / total status / port.
- Master TIF: WO Release, daftar LOP baru Go Live, dan alasan Drop.

Aturan utama:
- FBB total = PT2 + PT3 (tidak memasukkan baris blank).
- APC/Priority/Regular total Drop = Drop + Potensi Drop.
- APC & Priority OG Installation = Instalasi + Finish Installation.
- Regular tetap menampilkan Instalasi dan Finish Installation terpisah.
- Regular On Going memasukkan Blank Progress ke total, dan menampilkannya terpisah.
- Total APC diambil dari sheet Pivot APC; WO Released dari pivot APC di Control Data; WO Not Yet Released = selisih keduanya.
- Reason Drop hanya menghitung Status Project = DROP; tidak termasuk Potensi Drop.
- Satu LOP dapat masuk ke lebih dari satu reason jika dua kolom alasannya berbeda.
- Perbandingan tambahan Go Live menggunakan ID IHLD sebagai key.
"""

from __future__ import annotations

import os
import re
import sys
from datetime import datetime
from collections import Counter
from typing import Any, Dict, List, Optional, Tuple

from openpyxl import load_workbook
import pandas as pd


# ============================================================
# 1. CONFIGURATION - GANTI NAMA FILE DI BAGIAN INI SETIAP HARI
# ============================================================

HISTORY_FILE = "master data 200726 15-11-2026.xlsx"
TODAY_FILE = "master data 200726 16-11-2026.xlsx"

SHEET_CONTROL = "Control Data"
SHEET_MASTER = "Master TIF"
SHEET_APC_TOTAL = "Pivot APC"

REPORT_TIME = "15:00 WITA"
OUTPUT_DIR = "OUTPUT"
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "Text_Reporting_Perbandingan.txt")
VALIDATION_FILE = os.path.join(OUTPUT_DIR, "Validasi_Text_Reporting.txt")

# Standard pivot column regions based on Control Data 12-13 Nov 2026.
# Header row and Grand Total are searched dynamically inside each region.
PIVOTS = {
    "fbb":             {"cols": (1, 5),   "search_from": 14},
    "apc_stage":       {"cols": (8, 18),  "search_from": 14},
    "priority_lop":    {"cols": (21, 31), "search_from": 17},
    "priority_ports":  {"cols": (21, 31), "search_from": 43},
    "regular_lop":     {"cols": (36, 47), "search_from": 15},
    "regular_ports":   {"cols": (36, 47), "search_from": 39},
}

REASON_RULES = {
    "order duplicate": r"DUPLIK|ORDER DUP",
    "permit access rejected": r"PERIZIN|PERMIT",
    "drop DBAD": r"DROP DBAD|DBAD",
    "Drop Tsel": r"DROP TSEL|TSEL",
    "need redesign": r"REDESAIN|REDESIGN|RE-DESIGN",
    "Alpro issues": (
        r"ALPRO|DISTRIBUSI UNSPEC|DISTRIBUSI UNSPEK|DISTRIBUSI FULL|"
        r"FEEDER FULL|REDAMAN TINGGI"
    ),
    "Manpower Issues": r"MAN ?POWER|MANPOWER",
    "Material Issues": r"MATERIAL|OUT OF STOCK",
}


# ============================================================
# 2. GENERAL HELPERS
# ============================================================

def norm(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ").strip()).lower()


def clean(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").strip()
    return "" if text.lower() in {"", "none", "nan"} else text


def as_number(value: Any) -> int:
    if value is None or isinstance(value, bool):
        return 0
    if isinstance(value, (int, float)):
        return int(value)
    text = str(value).strip()
    if not text:
        return 0
    text = text.replace(",", "")
    try:
        return int(float(text))
    except (ValueError, TypeError):
        return 0


def fmt(value: Any) -> str:
    return f"{int(value):,}".replace(",", ".")


def pct(value: int, total: int) -> int:
    return round(value * 100 / total) if total else 0


def parse_date_from_filename(filename: str) -> datetime:
    match = re.search(r"(\d{2})-(\d{2})-(\d{4})", os.path.basename(filename))
    if not match:
        raise ValueError(
            f"Tidak dapat membaca tanggal dari nama file: {filename}. "
            "Gunakan format DD-MM-YYYY pada nama file."
        )
    day, month, year = map(int, match.groups())
    return datetime(year, month, day)


def date_label(dt: datetime) -> str:
    return dt.strftime("%d %b %Y")


def compare_count(today: int, history: int, history_label: str, unit: str = "LOP") -> str:
    diff = today - history
    if diff > 0:
        return f"▲ +{fmt(diff)} {unit} vs {history_label}"
    if diff < 0:
        return f"▼ -{fmt(abs(diff))} {unit} vs {history_label}"
    return f"= {history_label}"


def compare_percent(today: int, history: int, history_label: str) -> str:
    diff = today - history
    if diff > 0:
        return f"▲ +{diff} percentage point vs {history_label}"
    if diff < 0:
        return f"▼ {diff} percentage point vs {history_label}"
    return f"= {history_label}"


def open_book(path: str):
    if not os.path.isfile(path):
        raise FileNotFoundError(f"File tidak ditemukan: {path}")
    # keep_links=False mempercepat pembukaan workbook besar dengan external links.
    return load_workbook(path, read_only=True, data_only=True, keep_links=False)


def require_sheet(wb, name: str):
    if name not in wb.sheetnames:
        raise ValueError(f"Sheet '{name}' tidak ditemukan. Sheet tersedia: {wb.sheetnames}")
    return wb[name]


# ============================================================
# 3. CONTROL DATA PIVOT READER
# ============================================================

def locate_pivot(ws, cols: Tuple[int, int], search_from: int) -> Dict[str, Any]:
    """Baca area pivot secara berurutan agar cepat pada workbook read_only."""
    start_col, end_col = cols
    max_scan = min(ws.max_row, search_from + 15)
    matrix = list(ws.iter_rows(
        min_row=search_from, max_row=ws.max_row,
        min_col=start_col, max_col=end_col, values_only=True
    ))

    header_idx = None
    for idx, row_values in enumerate(matrix[:max_scan - search_from + 1]):
        first = norm(row_values[0] if row_values else None)
        headers_norm = [norm(v) for v in row_values]
        if first in {"row labels", "label baris"} and "grand total" in headers_norm:
            header_idx = idx
            break

    if header_idx is None:
        raise ValueError(
            f"Header pivot tidak ditemukan pada kolom {start_col}:{end_col}, mulai baris {search_from}."
        )

    header_row = search_from + header_idx
    header_values = matrix[header_idx]
    headers = {norm(value): idx for idx, value in enumerate(header_values) if norm(value)}

    total_idx = None
    for idx in range(header_idx + 1, len(matrix)):
        label = norm(matrix[idx][0] if matrix[idx] else None)
        if label in {"grand total", "total keseluruhan"}:
            total_idx = idx
            break

    if total_idx is None:
        raise ValueError(f"Grand Total tidak ditemukan untuk pivot header baris {header_row}.")

    total_row = search_from + total_idx
    totals = {
        header: as_number(matrix[total_idx][col_idx])
        for header, col_idx in headers.items()
    }
    details: Dict[str, Dict[str, int]] = {}
    for row_values in matrix[header_idx + 1:total_idx]:
        label = clean(row_values[0] if row_values else None)
        if not label:
            continue
        details[label] = {
            header: as_number(row_values[col_idx])
            for header, col_idx in headers.items()
        }

    return {
        "header_row": header_row,
        "total_row": total_row,
        "headers": headers,
        "totals": totals,
        "details": details,
    }


def read_control_data(filename: str) -> Dict[str, Any]:
    wb = open_book(filename)
    ws = require_sheet(wb, SHEET_CONTROL)
    data = {}
    for key, cfg in PIVOTS.items():
        data[key] = locate_pivot(ws, cfg["cols"], cfg["search_from"])

    # APC Released pivot is a two-level header; Grand Total is on row after header.
    # We identify the row dynamically using the category header and Count/Sum subheaders.
    apc_rows = list(ws.iter_rows(min_row=40, max_row=ws.max_row, min_col=8, max_col=16, values_only=True))
    apc_header = None
    apc_header_idx = None
    for idx, row_values in enumerate(apc_rows):
        values = [norm(v) for v in row_values]
        if "row labels" in values and "count of id ihld" in values and "sum of port" in values:
            apc_header = 40 + idx
            apc_header_idx = idx
            break
    if apc_header is None:
        raise ValueError("Header APC Released tidak ditemukan pada area H:P.")

    apc_total_row = None
    apc_total_values = None
    for idx in range(apc_header_idx + 1, len(apc_rows)):
        if norm(apc_rows[idx][0]) in {"grand total", "total keseluruhan"}:
            apc_total_row = 40 + idx
            apc_total_values = apc_rows[idx]
            break
    if apc_total_row is None:
        raise ValueError("Grand Total APC Released tidak ditemukan.")

    # Column mapping is based on the standard APC Released pivot:
    # I Drop LOP, J Drop Ports, K GL LOP, L GL Ports,
    # M On Going LOP, N On Going Ports, O Total LOP, P Total Ports.
    data["apc_released"] = {
        "header_row": apc_header,
        "total_row": apc_total_row,
        "drop_lop": as_number(apc_total_values[1]),
        "drop_ports": as_number(apc_total_values[2]),
        "go_live": as_number(apc_total_values[3]),
        "go_live_ports": as_number(apc_total_values[4]),
        "ongoing": as_number(apc_total_values[5]),
        "ongoing_ports": as_number(apc_total_values[6]),
        "total_lop": as_number(apc_total_values[7]),
        "total_ports": as_number(apc_total_values[8]),
    }
    wb.close()
    return data


def read_apc_total_pivot(filename: str) -> Dict[str, int]:
    """Baca total seluruh APC dari sheet Pivot APC (bukan Control Data)."""
    wb = open_book(filename)
    ws = require_sheet(wb, SHEET_APC_TOTAL)

    header_row = None
    count_col = None
    label_col = None

    # Temukan header berdasarkan Row Labels dan Count of iHLDLoPID.
    for row in range(1, min(ws.max_row, 30) + 1):
        row_values = [norm(ws.cell(row, col).value) for col in range(1, min(ws.max_column, 10) + 1)]
        for col, value in enumerate(row_values, start=1):
            if value in {"row labels", "label baris"}:
                label_col = col
            if "count of ihldlopid" in value or "count of id ihld" in value:
                count_col = col
        if label_col and count_col:
            header_row = row
            break

    if header_row is None:
        wb.close()
        raise ValueError("Header Count APC tidak ditemukan pada sheet Pivot APC.")

    grand_total_row = None
    for row in range(header_row + 1, ws.max_row + 1):
        label = norm(ws.cell(row, label_col).value)
        if label in {"grand total", "total keseluruhan"}:
            grand_total_row = row
            break

    if grand_total_row is None:
        wb.close()
        raise ValueError("Grand Total tidak ditemukan pada sheet Pivot APC.")

    total_apc = as_number(ws.cell(grand_total_row, count_col).value)
    total_ports = 0
    for col in range(count_col + 1, min(ws.max_column, count_col + 4) + 1):
        header = norm(ws.cell(header_row, col).value)
        if "sum of port" in header or "sum of portplan" in header:
            total_ports = as_number(ws.cell(grand_total_row, col).value)
            break

    wb.close()
    return {"total": total_apc, "total_ports": total_ports,
            "header_row": header_row, "grand_total_row": grand_total_row}


def apc_release_metrics(filename: str, control: Dict[str, Any]) -> Dict[str, int]:
    """Total APC dari Pivot APC; released/status/ports dari Control Data APC Released."""
    total_pivot = read_apc_total_pivot(filename)
    released_pivot = control["apc_released"]

    total_apc = total_pivot["total"]
    released = released_pivot["total_lop"]
    not_released = total_apc - released

    if not_released < 0:
        raise ValueError(
            f"WO Released ({released}) melebihi Total APC dari Pivot APC ({total_apc}) pada {filename}."
        )

    return {
        "total": total_apc,
        "released": released,
        "not_released": not_released,
        "total_ports": released_pivot["total_ports"],
        "go_live": released_pivot["go_live"],
        "go_live_ports": released_pivot["go_live_ports"],
        "ongoing": released_pivot["ongoing"],
        "ongoing_ports": released_pivot["ongoing_ports"],
        "drop": released_pivot["drop_lop"],
        "drop_ports": released_pivot["drop_ports"],
        "pivot_total_ports": total_pivot["total_ports"],
    }


def total_value(pivot: Dict[str, Any], aliases: List[str]) -> int:
    for alias in aliases:
        key = norm(alias)
        if key in pivot["totals"]:
            return pivot["totals"][key]
    return 0


def row_values(pivot: Dict[str, Any], aliases: List[str]) -> Dict[str, int]:
    for name, values in pivot["details"].items():
        if norm(name) in {norm(a) for a in aliases}:
            return values
    return {}


def stage_value(pivot: Dict[str, Any], aliases: List[str]) -> int:
    for alias in aliases:
        key = norm(alias)
        if key in pivot["totals"]:
            return pivot["totals"][key]
    return 0


# ============================================================
# 4. MASTER TIF READER / CLASSIFICATION
# ============================================================

MASTER_FIELDS = [
    "Unique_id", "ID IHLD", "IHLD FILTER", "NAMA LOP", "Port",
    "Status Project", "Detail Progress", "Kendala Deploy", "SUB KENDALA",
    "Tgl WO Release", "APC", "Priority Monthly",
]


def read_master_records(filename: str) -> List[Dict[str, Any]]:
    # pandas usecols jauh lebih cepat untuk Master TIF besar dibanding membaca
    # seluruh 67 kolom melalui iter_rows. openpyxl tetap dipakai untuk pivot.
    try:
        df = pd.read_excel(
            filename,
            sheet_name=SHEET_MASTER,
            usecols=lambda col: clean(col) in MASTER_FIELDS,
            engine="openpyxl",
        )
    except Exception as exc:
        raise RuntimeError(f"Gagal membaca Master TIF dari {filename}: {exc}") from exc

    missing = [field for field in MASTER_FIELDS if field not in df.columns]
    if missing:
        raise ValueError(f"Kolom Master TIF tidak ditemukan: {missing}")

    df = df[df["IHLD FILTER"].astype(str).str.strip().str.lower().eq("unik")].copy()
    df = df[df["ID IHLD"].notna()].copy()
    df["_key"] = df["ID IHLD"].astype(str).str.strip()
    return df.to_dict(orient="records")


def classify_record(record: Dict[str, Any]) -> str:
    if norm(record.get("APC")) == "apc":
        return "APC"
    if clean(record.get("Priority Monthly")):
        return "Priority"
    return "Regular"


def is_status(record: Dict[str, Any], status: str) -> bool:
    return norm(record.get("Status Project")) == norm(status)


def new_go_live_records(history_records: List[Dict[str, Any]], today_records: List[Dict[str, Any]]):
    history_by_key = {r["_key"]: r for r in history_records}
    added = []
    for record in today_records:
        if not is_status(record, "Go Live"):
            continue
        old = history_by_key.get(record["_key"])
        if old is None or not is_status(old, "Go Live"):
            added.append(record)
    return added


def unique_names(records: List[Dict[str, Any]]) -> List[str]:
    names = []
    seen = set()
    for record in records:
        name = clean(record.get("NAMA LOP"))
        key = record.get("_key")
        if name and key not in seen:
            names.append(name)
            seen.add(key)
    return names


def is_released(record: Dict[str, Any]) -> bool:
    """True jika Tgl WO Release berisi tanggal/nilai valid, bukan placeholder error."""
    value = record.get("Tgl WO Release")
    text = clean(value).strip().lower()
    if text in {"", "none", "nan", "nat", "#n/a", "n/a", "#value!", "-", "0"}:
        return False
    return True


def numeric_port(record: Dict[str, Any]) -> int:
    return as_number(record.get("Port"))


def drop_reason_counts(records: List[Dict[str, Any]], category: str, status: str) -> Dict[str, Any]:
    """Hitung reason untuk satu kategori dan satu status (DROP/POTENSI DROP).

    Blank hanya jika Kendala Deploy DAN SUB KENDALA sama-sama kosong.
    Other / Unclassified untuk teks terisi yang tidak cocok dengan regex mana pun.
    Set per ID IHLD menjaga hitungan unik; kategori reason dapat overlap.
    """
    selected = [
        r for r in records
        if classify_record(r) == category and is_status(r, status)
    ]
    reason_matches: Dict[str, set] = {label: set() for label in REASON_RULES}
    blank_ids, other_ids, any_match = set(), set(), set()
    all_ids = {r["_key"] for r in selected}

    for record in selected:
        key = record["_key"]
        kendala = clean(record.get("Kendala Deploy"))
        sub = clean(record.get("SUB KENDALA"))
        text = f"{kendala} | {sub}".upper()
        if not kendala and not sub:
            blank_ids.add(key)
            continue
        matched = False
        for label, pattern in REASON_RULES.items():
            if re.search(pattern, text, flags=re.IGNORECASE):
                reason_matches[label].add(key)
                any_match.add(key)
                matched = True
        if not matched:
            other_ids.add(key)

    counts = {label: len(ids) for label, ids in reason_matches.items() if ids}
    return {
        "status": status,
        "counts": counts,
        "status_unique": len(all_ids),
        "classified_unique": len(any_match),
        "blank": len(blank_ids),
        "other": len(other_ids),
        "category_sum": sum(counts.values()),
        "overlap": max(0, sum(counts.values()) - len(any_match)),
    }


def format_reasons(result: Dict[str, Any]) -> str:
    parts = [f"{fmt(count)} {label}" for label, count in result["counts"].items()]
    if result["other"]:
        parts.append(f"{fmt(result['other'])} Other / Unclassified")
    if result["blank"]:
        parts.append(f"{fmt(result['blank'])} Blank")
    return ", ".join(parts) if parts else "Belum ada data reason"


def append_propose_drop_reasons(lines: List[str], category: str,
                                propose_drop: int, today: int, history: int,
                                history_label: str,
                                reasons: Dict[str, Dict[str, Any]]) -> None:
    lines.append(
        f"* Propose Drop: {fmt(propose_drop)} LOP "
        f"({compare_count(today, history, history_label)})"
    )
    for status in ["DROP", "POTENSI DROP"]:
        result = reasons[category][status]
        lines.append(f"  - {status}: {format_reasons(result)}")


# ============================================================
# 5. METRIC EXTRACTION FROM CONTROL DATA
# ============================================================

def extract_metrics(control: Dict[str, Any]) -> Dict[str, Any]:
    # FBB: exclude (blank) by summing PT2 + PT3, not Grand Total.
    fbb_pivot = control["fbb"]
    pt2 = total_value(fbb_pivot, ["PT2"])
    pt3 = total_value(fbb_pivot, ["PT3"])
    fbb_rows = fbb_pivot["details"]

    def fbb_status_by_pt(label: str, pt_key: str) -> int:
        values = row_values(fbb_pivot, [label])
        return as_number(values.get(norm(pt_key), 0))

    def fbb_status(label: str) -> int:
        return fbb_status_by_pt(label, "PT2") + fbb_status_by_pt(label, "PT3")

    fbb = {
        "pt2": pt2,
        "pt3": pt3,
        "total": pt2 + pt3,
        "go_live": fbb_status("Go Live"),
        "ongoing": fbb_status("On Going"),
        "drop": fbb_status("Drop"),
        "pt2_go_live": fbb_status_by_pt("Go Live", "PT2"),
        "pt2_ongoing": fbb_status_by_pt("On Going", "PT2"),
        "pt2_drop": fbb_status_by_pt("Drop", "PT2"),
        "pt3_go_live": fbb_status_by_pt("Go Live", "PT3"),
        "pt3_ongoing": fbb_status_by_pt("On Going", "PT3"),
        "pt3_drop": fbb_status_by_pt("Drop", "PT3"),
    }
    fbb["percentage"] = pct(fbb["go_live"], fbb["total"])
    fbb["pt2_percentage"] = pct(fbb["go_live"], pt2)
    fbb["pt3_percentage"] = pct(fbb["go_live"], pt3)

    # APC stages: H:R Grand Total row. Drop total combines Drop + Potensi Drop.
    apc_p = control["apc_stage"]
    apc = {
        "potential_drop": stage_value(apc_p, ["00. potensi drop"]),
        "drop_only": stage_value(apc_p, ["00. drop"]),
        "survey": stage_value(apc_p, ["01. survey & anwizing"]),
        "drm": stage_value(apc_p, ["02. drm"]),
        "preparing": stage_value(apc_p, ["03. preparing"]),
        "material": stage_value(apc_p, ["04. matdel"]),
        "installation": stage_value(apc_p, ["05. instalasi"]),
        "finish": stage_value(apc_p, ["06. finish install"]),
        "go_live": stage_value(apc_p, ["07. golive"]),
    }
    apc["propose_drop"] = apc["potential_drop"] + apc["drop_only"]
    apc["og_installation"] = apc["installation"] + apc["finish"]

    # Priority LOP.
    pr_p = control["priority_lop"]
    priority = {
        "total": total_value(pr_p, ["grand total"]),
        "drop_only": stage_value(pr_p, ["00. drop"]),
        "potential_drop": stage_value(pr_p, ["00. potensi drop"]),
        "survey": stage_value(pr_p, ["01. survey & anwizing"]),
        "drm": stage_value(pr_p, ["02. drm"]),
        "preparing": stage_value(pr_p, ["03. preparing"]),
        "material": stage_value(pr_p, ["04. matdel"]),
        "installation": stage_value(pr_p, ["05. instalasi"]),
        "finish": stage_value(pr_p, ["06. finish install"]),
        "go_live": stage_value(pr_p, ["07. golive"]),
    }
    priority["propose_drop"] = priority["drop_only"] + priority["potential_drop"]
    priority["og_installation"] = priority["installation"] + priority["finish"]
    priority["ongoing"] = sum(priority[k] for k in ["survey", "drm", "preparing", "material", "og_installation"])
    priority["not_yet_go_live"] = priority["total"] - priority["go_live"]
    priority["percentage"] = pct(priority["go_live"], priority["total"])

    # Priority Ports pivot.
    pr_ports = control["priority_ports"]
    priority["total_ports"] = total_value(pr_ports, ["grand total"])
    priority["go_live_ports"] = stage_value(pr_ports, ["07. golive"])
    priority["ongoing_ports"] = sum(
        stage_value(pr_ports, [alias]) for alias in [
            "01. survey & anwizing", "02. drm", "03. preparing",
            "04. matdel", "05. instalasi", "06. finish install"
        ]
    )

    # Regular LOP and Ports. Blank is retained and shown separately.
    rg_p = control["regular_lop"]
    regular = {
        "total": total_value(rg_p, ["grand total"]),
        "drop_only": stage_value(rg_p, ["00. drop"]),
        "potential_drop": stage_value(rg_p, ["00. potensi drop"]),
        "survey": stage_value(rg_p, ["01. survey & anwizing"]),
        "drm": stage_value(rg_p, ["02. drm"]),
        "preparing": stage_value(rg_p, ["03. preparing"]),
        "material": stage_value(rg_p, ["04. matdel"]),
        "installation": stage_value(rg_p, ["05. instalasi"]),
        "finish": stage_value(rg_p, ["06. finish install"]),
        "go_live": stage_value(rg_p, ["07. golive"]),
        "blank": stage_value(rg_p, ["(blank)"]),
    }
    regular["propose_drop"] = regular["drop_only"] + regular["potential_drop"]
    regular["ongoing_stages"] = sum(regular[k] for k in ["survey", "drm", "preparing", "material", "installation", "finish"])
    regular["ongoing"] = regular["ongoing_stages"] + regular["blank"]
    regular["not_yet_go_live"] = regular["total"] - regular["go_live"]
    regular["percentage"] = pct(regular["go_live"], regular["total"])

    rg_ports = control["regular_ports"]
    regular["total_ports"] = total_value(rg_ports, ["grand total"])
    regular["go_live_ports"] = stage_value(rg_ports, ["07. golive"])
    regular["ongoing_ports_stages"] = sum(
        stage_value(rg_ports, [alias]) for alias in [
            "01. survey & anwizing", "02. drm", "03. preparing",
            "04. matdel", "05. instalasi", "06. finish install"
        ]
    )
    regular["blank_ports"] = stage_value(rg_ports, ["(blank)"])
    regular["ongoing_ports"] = regular["ongoing_ports_stages"] + regular["blank_ports"]

    return {"fbb": fbb, "apc": apc, "priority": priority, "regular": regular}


# ============================================================
# 6. REPORT FORMATTERS
# ============================================================

def stage_line(label: str, today: int, history: int, history_label: str) -> str:
    return f"* {label}: {fmt(today)} LOP ({compare_count(today, history, history_label)})"


def names_line(names: List[str]) -> str:
    if not names:
        return "There are no additional LOP Go Live since the previous report"
    return ", ".join(names)


def build_report(
    hist_metrics: Dict[str, Any], today_metrics: Dict[str, Any],
    hist_apc_release: Dict[str, int], today_apc_release: Dict[str, int],
    hist_records: List[Dict[str, Any]], today_records: List[Dict[str, Any]],
    report_dt: datetime, history_dt: datetime,
) -> Tuple[str, str]:
    hlabel = date_label(history_dt)
    report_label = date_label(report_dt)

    # Additional GL by ID IHLD, then partition using Master TIF category rule.
    new_gl = new_go_live_records(hist_records, today_records)
    new_by_category = {
        category: [r for r in new_gl if classify_record(r) == category]
        for category in ["APC", "Priority", "Regular"]
    }

    # Reason DROP dan POTENSI DROP dipisahkan untuk setiap kategori.
    reasons = {
        category: {
            status: drop_reason_counts(today_records, category, status)
            for status in ["DROP", "POTENSI DROP"]
        }
        for category in ["APC", "Priority", "Regular"]
    }

    f, fh = today_metrics["fbb"], hist_metrics["fbb"]
    a, ah = today_apc_release, hist_apc_release
    apc_stage, apc_stage_h = today_metrics["apc"], hist_metrics["apc"]
    p, ph = today_metrics["priority"], hist_metrics["priority"]
    r, rh = today_metrics["regular"], hist_metrics["regular"]

    lines: List[str] = []
    lines += [
        "*Update Progress Deployment FBB Area-4 PAMASUKA*",
        f"Status on {report_label} {REPORT_TIME}",
        "========================================================",
        "",
        "*Total WO FBB Pamasuka: Go Live | On going | Drop | % Go Live*",
        (
            f"{fmt(f['total'])} LOP : {fmt(f['go_live'])} "
            f"({compare_count(f['go_live'], fh['go_live'], hlabel)}) | "
            f"{fmt(f['ongoing'])} ({compare_count(f['ongoing'], fh['ongoing'], hlabel)}) | "
            f"{fmt(f['drop'])} ({compare_count(f['drop'], fh['drop'], hlabel)}) | "
            f"{f['percentage']}% ({compare_percent(f['percentage'], fh['percentage'], hlabel)})"
        ),
        "",
        "*PT2 WO FBB Pamasuka*",
        "*Total WO PT2: Go Live | On going | Drop | % Go Live*",
        (
            f"{fmt(f['pt2'])} LOP : {fmt(f['pt2_go_live'])} "
            f"({compare_count(f['pt2_go_live'], fh['pt2_go_live'], hlabel)}) | "
            f"{fmt(f['pt2_ongoing'])} ({compare_count(f['pt2_ongoing'], fh['pt2_ongoing'], hlabel)}) | "
            f"{fmt(f['pt2_drop'])} ({compare_count(f['pt2_drop'], fh['pt2_drop'], hlabel)}) | "
            f"{pct(f['pt2_go_live'], f['pt2'])}%"
        ),
        "",
        "*PT3 WO FBB Pamasuka*",
        "*Total WO PT3: Go Live | On going | Drop | % Go Live*",
        (
            f"{fmt(f['pt3'])} LOP : {fmt(f['pt3_go_live'])} "
            f"({compare_count(f['pt3_go_live'], fh['pt3_go_live'], hlabel)}) | "
            f"{fmt(f['pt3_ongoing'])} ({compare_count(f['pt3_ongoing'], fh['pt3_ongoing'], hlabel)}) | "
            f"{fmt(f['pt3_drop'])} ({compare_count(f['pt3_drop'], fh['pt3_drop'], hlabel)}) | "
            f"{pct(f['pt3_go_live'], f['pt3'])}%"
        ),
    ]

    # APC released section - all metrics are based on Master TIF release date.
    lines += [
        "",
        "*Total LOP APC : WO Released | WO NY Released | %*",
        f"{fmt(a['total'])} : {fmt(a['released'])} | {fmt(a['not_released'])} | {pct(a['released'], a['total'])}%",
        "",
        "A. HIGHLIGHT SUMMARY",
        f"1. Total LOP: {fmt(a['total'])} ({compare_count(a['total'], ah['total'], hlabel)})",
        f"2. WO Released: {fmt(a['released'])} LOP ({compare_count(a['released'], ah['released'], hlabel)}) (dari pivot APC Control Data)",
        f"3. WO Not Yet Released: {fmt(a['not_released'])} LOP ({compare_count(a['not_released'], ah['not_released'], hlabel)}) (Total Pivot APC - WO Released Control Data)",
        f"4. Go Live : {fmt(a['go_live'])} LOP ({compare_count(a['go_live'], ah['go_live'], hlabel)})",
        "",
        f"👉 Note: Need Support for {fmt(a['not_released'])} LOP Not Yet WO Release",
        f"% Released: {pct(a['released'], a['total'])}% ({compare_percent(pct(a['released'], a['total']), pct(ah['released'], ah['total']), hlabel)})",
        "",
        f"B. STATUS WO RELEASED: {fmt(a['released'])} LOP / {fmt(a['total_ports'])} Ports",
        "",
        (
            f"1. GO LIVE: {fmt(a['go_live'])} LOP / {fmt(a['go_live_ports'])} Ports, "
            f"{pct(a['go_live'], a['released'])}% ({compare_count(a['go_live'], ah['go_live'], hlabel)})"
        ),
        (
            f"(Progress: There are {len(new_by_category['APC'])} additional LOP Go Live vs {hlabel}: "
            f"*{names_line(unique_names(new_by_category['APC']))}*)"
        ),
        "2. ON GOING: "
        f"{fmt(a['ongoing'])} LOP / {fmt(a['ongoing_ports'])} Ports, "
        f"{pct(a['ongoing'], a['released'])}% ({compare_count(a['ongoing'], ah['ongoing'], hlabel)})",
        stage_line("Survey & Anwizing", apc_stage["survey"], apc_stage_h["survey"], hlabel),
        stage_line("DRM", apc_stage["drm"], apc_stage_h["drm"], hlabel),
        stage_line("Preparing", apc_stage["preparing"], apc_stage_h["preparing"], hlabel),
        stage_line("Material OG Delivery", apc_stage["material"], apc_stage_h["material"], hlabel),
        stage_line("OG Installation", apc_stage["og_installation"], apc_stage_h["og_installation"], hlabel),
    ]
    apc_stage_sum = sum(apc_stage[k] for k in ["survey", "drm", "preparing", "material", "og_installation"])
    apc_other_ongoing = max(0, a["ongoing"] - apc_stage_sum)
    if apc_other_ongoing:
        lines.append(f"* Other / Unclassified Progress: {fmt(apc_other_ongoing)} LOP (selisih pivot stage dengan Master TIF WO Released)")
    append_propose_drop_reasons(
        lines, "APC", apc_stage["propose_drop"], apc_stage["propose_drop"],
        apc_stage_h["propose_drop"], hlabel, reasons
    )

    # PRIORITY
    lines += [
        "",
        "*Total LOP Priority November 2026 : Go Live | NY Go Live | %*",
        f"{fmt(p['total'])} : {fmt(p['go_live'])} | {fmt(p['not_yet_go_live'])} | {p['percentage']}%",
        "",
        "A. HIGHLIGHT SUMMARY",
        f"1. Total LOP: {fmt(p['total'])} ({compare_count(p['total'], ph['total'], hlabel)})",
        f"2. Go Live: {fmt(p['go_live'])} LOP ({compare_count(p['go_live'], ph['go_live'], hlabel)})",
        f"3. Not Yet Go Live: {fmt(p['not_yet_go_live'])} LOP ({compare_count(p['not_yet_go_live'], ph['not_yet_go_live'], hlabel)})",
        "",
        "👉 Note: 23 LOP intersection with LOP APC, already takeout (sesuaikan jika angka intersection berubah).",
        f"   Total % Go Live : {p['percentage']}%",
        "",
        f"B. STATUS LOP PRIORITY OCT 2026 : {fmt(p['total'])} LOP / {fmt(p['total_ports'])} Ports "
        f"({compare_count(p['total'], ph['total'], hlabel)})",
        (
            f"1. GO LIVE: {fmt(p['go_live'])} LOP / {fmt(p['go_live_ports'])} Ports, "
            f"{pct(p['go_live'], p['total'])}% ({compare_count(p['go_live'], ph['go_live'], hlabel)})"
        ),
        (
            f"(Progress: There are {len(new_by_category['Priority'])} additional LOP Go Live vs {hlabel}: "
            f"*{names_line(unique_names(new_by_category['Priority']))}*)"
        ),
        f"2. ON GOING: {fmt(p['ongoing'])} LOP / {fmt(p['ongoing_ports'])} Ports, "
        f"{pct(p['ongoing'], p['total'])}% ({compare_count(p['ongoing'], ph['ongoing'], hlabel)})",
        stage_line("Survey & Anwizing", p["survey"], ph["survey"], hlabel),
        stage_line("DRM", p["drm"], ph["drm"], hlabel),
        stage_line("Preparing", p["preparing"], ph["preparing"], hlabel),
        stage_line("Material OG Delivery", p["material"], ph["material"], hlabel),
        stage_line("OG Installation", p["og_installation"], ph["og_installation"], hlabel),
        f"* Propose Drop: {fmt(p['propose_drop'])} LOP ({compare_count(p['propose_drop'], ph['propose_drop'], hlabel)})",
        f"  - DROP: {format_reasons(reasons['Priority']['DROP'])}",
        f"  - POTENSI DROP: {format_reasons(reasons['Priority']['POTENSI DROP'])}",
    ]

    # REGULAR
    lines += [
        "",
        "*Total LOP Reguler : Go Live | NY Go Live | % Go Live*",
        f"{fmt(r['total'])} : {fmt(r['go_live'])} | {fmt(r['not_yet_go_live'])} | {r['percentage']}%",
        "",
        "A. HIGHLIGHT SUMMARY",
        f"1. Total LOP: {fmt(r['total'])} ({compare_count(r['total'], rh['total'], hlabel)})",
        f"2. Go Live: {fmt(r['go_live'])} LOP ({compare_count(r['go_live'], rh['go_live'], hlabel)})",
        f"3. Not Yet Go Live: {fmt(r['not_yet_go_live'])} LOP ({compare_count(r['not_yet_go_live'], rh['not_yet_go_live'], hlabel)})",
        "",
        "👉 Note: Need Support additional Manpower to support not only on LOP APC & LOP Priority but also LOP Reguler besides trm for MBB Target",
        f"% Go Live: {r['percentage']}% ({compare_percent(r['percentage'], rh['percentage'], hlabel)})",
        "",
        f"B. STATUS LOP REGULER: {fmt(r['total'])} LOP / {fmt(r['total_ports'])} Ports "
        f"({compare_count(r['total'], rh['total'], hlabel)})",
        (
            f"1. GO LIVE: {fmt(r['go_live'])} LOP / {fmt(r['go_live_ports'])} Ports, "
            f"{r['percentage']}% ({compare_count(r['go_live'], rh['go_live'], hlabel)})"
        ),
        (
            f"(Progress: There are {len(new_by_category['Regular'])} additional LOP Go Live vs {hlabel}: "
            f"*{names_line(unique_names(new_by_category['Regular']))}*)"
        ),
        f"2. ON GOING: {fmt(r['ongoing'])} LOP / {fmt(r['ongoing_ports'])} Ports, "
        f"{pct(r['ongoing'], r['total'])}% ({compare_count(r['ongoing'], rh['ongoing'], hlabel)})",
        stage_line("Survey & Anwizing", r["survey"], rh["survey"], hlabel),
        stage_line("DRM", r["drm"], rh["drm"], hlabel),
        stage_line("Preparing", r["preparing"], rh["preparing"], hlabel),
        stage_line("Material OG Delivery", r["material"], rh["material"], hlabel),
        stage_line("OG Installation", r["installation"], rh["installation"], hlabel),
        stage_line("Finish Installation", r["finish"], rh["finish"], hlabel),
    ]
    if r["blank"]:
        lines.append(
            f"* Unclassified / Blank Progress: {fmt(r['blank'])} LOP "
            f"({compare_count(r['blank'], rh['blank'], hlabel)})"
        )
    lines += [
        f"* Propose Drop: {fmt(r['propose_drop'])} LOP ({compare_count(r['propose_drop'], rh['propose_drop'], hlabel)})",
        f"  - DROP: {format_reasons(reasons['Regular']['DROP'])}",
        f"  - POTENSI DROP: {format_reasons(reasons['Regular']['POTENSI DROP'])}",
        "",
        "VALIDASI OTOMATIS:",
        f"- Master TIF: {len(new_gl)} LOP baru berstatus Go Live berdasarkan ID IHLD.",
        f"- Rincian tambahan Go Live: APC {len(new_by_category['APC'])}, Priority {len(new_by_category['Priority'])}, Regular {len(new_by_category['Regular'])} LOP.",
        "- Klasifikasi: APC dari APC='APC'; Priority dari Priority Monthly terisi dan bukan APC; Regular selain keduanya.",
        "- Total Drop/Propose Drop pada APC, Priority, dan Regular = kolom DROP + POTENSI DROP dari Control Data.",
        "- OG Installation APC dan Priority = INSTALASI + FINISH INSTALL; delta dihitung dari hasil gabungan.",
        "- Reason DROP dan POTENSI DROP dihitung terpisah berdasarkan Status Project.",
        "- Jika Kendala Deploy dan SUB KENDALA sama-sama kosong, reason ditampilkan sebagai Blank.",
        "- Jika teks alasan terisi tetapi tidak cocok regex, masuk Other / Unclassified; satu LOP dapat overlap antar-reason.",
        f"- APC Total: {a['total']} LOP dari sheet Pivot APC.",
        f"- APC WO Released: {a['released']} LOP dari pivot APC sheet Control Data.",
        f"- APC WO Not Yet Released: {a['not_released']} LOP = Pivot APC total - Control Data WO Released.",
        f"- APC Master TIF unik: {a['total']} LOP; APC stage pivot Control Data: {sum(today_metrics['apc'][k] for k in ['propose_drop', 'survey', 'drm', 'preparing', 'material', 'og_installation', 'go_live'])} LOP (cek silang status).",
        f"- Regular Blank Progress: {r['blank']} LOP dan {r['blank_ports']} Ports; dimasukkan ke total On Going.",
    ]

    # Detailed validation report, separate from WhatsApp report.
    validation = build_validation(
        hist_metrics, today_metrics, hist_apc_release, today_apc_release,
        hist_records, today_records, new_gl, reasons, hlabel, report_label
    )
    return "\n".join(lines), validation


# ============================================================
# 7. VALIDATION SUMMARY
# ============================================================

def build_validation(hist_metrics, today_metrics, hist_apc, today_apc,
                     hist_records, today_records, new_gl, reasons, hlabel, today_label):
    lines = [
        f"VALIDASI TEXT REPORTING - {today_label} dibanding {hlabel}",
        "=" * 72,
        "",
        "1. CONTROL DATA PIVOT",
    ]
    for label, key in [("FBB", "fbb"), ("Priority", "priority"), ("Regular", "regular")]:
        h = hist_metrics[key]
        t = today_metrics[key]
        lines.append(f"{label}: HISTORY total={h['total']:,}; TODAY total={t['total']:,}.")
    lines += [
        "",
        "2. APC TOTAL & WO RELEASED - PIVOT APC + CONTROL DATA",
        f"HISTORY: total APC Pivot APC={hist_apc['total']}, WO Released Control Data={hist_apc['released']}, not released={hist_apc['not_released']}, released ports={hist_apc['total_ports']}.",
        f"TODAY: total APC Pivot APC={today_apc['total']}, WO Released Control Data={today_apc['released']}, not released={today_apc['not_released']}, released ports={today_apc['total_ports']}.",
        "",
        "3. MASTER TIF CATEGORY COUNTS (IHLD FILTER=unik)",
    ]
    for category in ["APC", "Priority", "Regular"]:
        count = sum(1 for rec in today_records if classify_record(rec) == category)
        lines.append(f"{category}: {count} LOP.")
    lines += [
        "Catatan: kategori Master TIF menggunakan aturan yang diminta. Angka laporan utama tetap mengambil pivot Control Data; jika berbeda, jangan dipaksa sama tanpa mengubah filter sumber.",
        "",
        "4. NEW GO LIVE",
        f"Total tambahan Go Live berdasarkan ID IHLD: {len(new_gl)} LOP.",
    ]
    for category in ["APC", "Priority", "Regular"]:
        count = sum(1 for rec in new_gl if classify_record(rec) == category)
        lines.append(f"{category}: {count} LOP.")
    lines += ["", "5. REASON DROP & POTENSI DROP (TERPISAH)"]
    for category, status_results in reasons.items():
        for status, result in status_results.items():
            lines.append(
                f"{category} / {status}: unique={result['status_unique']}; "
                f"matched unique={result['classified_unique']}; blank={result['blank']}; "
                f"other={result['other']}; sum reason categories={result['category_sum']}; "
                f"overlap={result['overlap']}."
            )
            for label, count in result["counts"].items():
                lines.append(f"  - {label}: {count}")
    lines += [
        "",
        "PENTING: Total APC berasal dari sheet Pivot APC. WO Released berasal dari pivot APC di Control Data. WO Not Yet Released = Total APC Pivot APC - WO Released Control Data.",
        "Total Propose Drop pada laporan = DROP + POTENSI DROP dari pivot. Breakdown reason Master TIF dipisahkan berdasarkan Status Project=DROP dan Status Project=POTENSI DROP.",
        "Blank berarti Kendala Deploy dan SUB KENDALA sama-sama kosong. Other / Unclassified berarti minimal satu kolom terisi tetapi tidak cocok dengan regex reason.",
        "Reason categories dapat overlap; satu LOP dapat terhitung di lebih dari satu kategori reason. Semua hitungan berdasarkan ID IHLD unik.",
        "",
        "6. BALANCE CHECK",
    ]
    for category in ["priority", "regular"]:
        m = today_metrics[category]
        name = category.upper()
        if category == "priority":
            stage_sum = m["go_live"] + m["ongoing"] + m["propose_drop"]
            lines.append(f"{name}: GL {m['go_live']} + On Going {m['ongoing']} + Drop/Propose {m['propose_drop']} = {stage_sum}; Total {m['total']}.")
        else:
            stage_sum = m["go_live"] + m["ongoing"] + m["propose_drop"]
            lines.append(f"{name}: GL {m['go_live']} + On Going incl. blank {m['ongoing']} + Drop/Propose {m['propose_drop']} = {stage_sum}; Total {m['total']}.")
    apc_status_sum = today_apc['go_live'] + today_apc['ongoing'] + today_apc['drop']
    lines.append(f"APC Released Control Data: GL {today_apc['go_live']} + On Going {today_apc['ongoing']} + Drop {today_apc['drop']} = {apc_status_sum}; Released total {today_apc['released']}.")
    lines.append(f"APC total source Pivot APC={today_apc['total']}; WO Released source Control Data={today_apc['released']}; WO Not Yet Released={today_apc['not_released']}.")
    lines.append(f"APC Control Data stage: Propose Drop {today_metrics['apc']['propose_drop']} + On Going stages {sum(today_metrics['apc'][k] for k in ['survey','drm','preparing','material','og_installation'])} + Go Live {today_metrics['apc']['go_live']} = {sum(today_metrics['apc'][k] for k in ['propose_drop','survey','drm','preparing','material','og_installation','go_live'])}.")
    lines.append("Catatan APC: jika komposisi status Master TIF dan Control Data berbeda, keduanya ditampilkan sebagai sumber berbeda; jangan dipaksa sama.")
    return "\n".join(lines)


# ============================================================
# 8. MAIN
# ============================================================

def main():
    history_dt = parse_date_from_filename(HISTORY_FILE)
    report_dt = parse_date_from_filename(TODAY_FILE)

    print("=" * 72)
    print("FBB AREA-4 PAMASUKA - TEXT REPORTING FULL REV FINAL 6.0")
    print("=" * 72)
    print(f"HISTORY : {HISTORY_FILE}")
    print(f"TODAY   : {TODAY_FILE}")

    print("\n[1/4] Membaca Control Data HISTORY...")
    history_control = read_control_data(HISTORY_FILE)
    print("[2/4] Membaca Control Data TODAY...")
    today_control = read_control_data(TODAY_FILE)

    print("[3/4] Menghitung metrik Master TIF HISTORY & TODAY...")
    history_records = read_master_records(HISTORY_FILE)
    today_records = read_master_records(TODAY_FILE)

    history_metrics = extract_metrics(history_control)
    today_metrics = extract_metrics(today_control)
    history_apc_release = apc_release_metrics(HISTORY_FILE, history_control)
    today_apc_release = apc_release_metrics(TODAY_FILE, today_control)

    print("[4/4] Membuat laporan dan validasi...")
    report, validation = build_report(
        history_metrics, today_metrics,
        history_apc_release, today_apc_release,
        history_records, today_records,
        report_dt, history_dt,
    )

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(report)
    with open(VALIDATION_FILE, "w", encoding="utf-8") as f:
        f.write(validation)

    print("\n" + "=" * 72)
    print("SELESAI")
    print(f"Report    : {OUTPUT_FILE}")
    print(f"Validasi  : {VALIDATION_FILE}")
    print("=" * 72)
    print("\nPreview report:\n")
    print(report)


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\nERROR: {exc}")
        print("Periksa nama file, nama sheet, dan struktur pivot Control Data.")
        input("Tekan ENTER untuk keluar...")
        sys.exit(1)
