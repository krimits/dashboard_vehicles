# -*- coding: utf-8 -*-
from __future__ import annotations

import argparse
import hmac
import json
import os
import re
from collections import Counter, defaultdict
from datetime import date, datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any
from urllib.parse import unquote, urlparse

import pandas as pd


DEFAULT_HOST = os.getenv("HOST", "0.0.0.0")
DEFAULT_PORT = int(os.getenv("PORT", "8000"))
EXCEL_PREFIX = "\u0394\u0395\u039b\u03a4\u0399\u039f \u0395\u039d\u0397\u039c\u0395\u03a1\u03a9\u03a3\u0397\u03a3"
LOCK_FILE_PREFIX = "~$"
POLL_INTERVAL_MS = 5000
ALLOWED_UPLOAD_EXTENSIONS = {".xls", ".xlsx", ".xlsm"}
DEFAULT_STORAGE_DIRNAME = "dashboard_data"
ACTIVE_WORKBOOK_METADATA = "active_workbook.json"
DEFAULT_ADMIN_SECRET_ENV = "DASHBOARD_ADMIN_SECRET"
MAX_UPLOAD_BYTES = 15 * 1024 * 1024
HEADER_MARKERS = {
    "\u039a\u0391\u03a4\u0397\u0393\u039f\u03a1\u0399\u0391 \u039f\u03a7\u0397\u039c.",
    "\u039a\u0391\u03a4\u0397\u0393\u039f\u03a1\u0399\u0391 \u039f\u03a7\u0397\u039c",
}
EXPLICIT_EXCLUSION_NOTES = {
    "\u0391\u03a0\u039f\u03a3\u03a5\u03a1\u03a3\u0397",
    "\u0391\u03a3\u03a5\u039c\u03a6\u039f\u03a1\u039f \u0395\u03a0\u0399\u03a3\u039a\u0395\u03a5\u0397\u03a3",
    "\u03a0\u03a1\u039f\u03a3 \u0391\u03a0\u039f\u03a3\u03a5\u03a1\u03a3\u0397",
    "\u03a0\u0391\u03a1\u0391\u03a7\u03a9\u03a1\u0397\u03a3\u0397",
}
INFERRED_EXCLUSION_NOTES = {
    "\u0391\u03a0\u039f\u03a3\u03a5\u03a1\u03a3\u0397",
    "\u0391\u03a3\u03a5\u039c\u03a6\u039f\u03a1\u039f \u0395\u03a0\u0399\u03a3\u039a\u0395\u03a5\u0397\u03a3",
    "\u03a0\u03a1\u039f\u03a3 \u0391\u03a0\u039f\u03a3\u03a5\u03a1\u03a3\u0397",
    "\u03a0\u0391\u03a1\u0391\u03a7\u03a9\u03a1\u0397\u03a3\u0397",
    "\u0391\u039d\u0391\u039c\u039f\u039d\u0397 \u0395\u039e. \u03a3\u03a5\u039d\u0395\u03a1\u0393\u0395\u0399\u039f",
    "\u0391\u039d\u0391\u039c\u039f\u039d\u0397 \u0395\u03a0\u0399\u03a3\u039a\u0395\u03a5\u0397\u03a3",
    "\u03a3\u03a5\u039d\u0395\u03a1\u0393\u0395\u0399\u039f",
    "\u0395\u039e\u03a9\u03a4\u0395\u03a1\u0399\u039a\u039f \u03a3\u03a5\u039d\u0395\u03a1\u0393\u0395\u0399\u039f",
}
MANAGEMENT_CATEGORY_ALIASES = {
    "\u039c\u0399\u039a\u03a1\u0391 2\u03a4 (\u0397\u039b\u0395\u039a\u03a4\u03a1\u0399\u039a\u0391)": "\u0397\u039b\u0395\u039a\u03a4\u03a1\u0399\u039a\u0391 \u039c\u0399\u039a\u03a1\u0391",
    "\u039c\u0399\u039a\u03a1\u0391 4\u03a4 (\u039c\u03a5\u039b\u039f\u0399)": "\u039c\u0399\u039a\u03a1\u0391 4\u03a4 (\u039c\u03a5\u039b\u039f\u0399) (IVECO)",
    "\u039d\u0395\u0391 \u039c\u0399\u039a\u03a1\u0391 2\u03a4 (\u03a0\u03a1\u0395\u03a3\u03a3\u0391\u039a\u0399\u0391)": "\u039d\u0395\u0391 \u039c\u0399\u039a\u03a1\u0391 2\u03a4 (\u03a0\u03a1\u0395\u03a3\u03a3\u0391\u039a\u0399\u0391)",
    "\u03a0\u0391\u039b\u0399\u0391 5\u03a4 (\u039c\u03a5\u039b\u039f\u0399)": "\u03a0\u0391\u039b\u0399\u0391 \u039c\u03a5\u039b\u039f\u0399 5\u03a4",
    "\u039d\u0395\u0391 5\u03a4 (\u039c\u03a5\u039b\u039f\u0399)": "\u039d\u0395\u0391 \u039c\u03a5\u039b\u039f\u0399 5\u03a4 (\u039c\u0391\u039d)",
    "\u03a0\u03a1\u0395\u03a3\u0395\u03a3 12\u03a4": "\u03a0\u03a1\u0395\u03a3\u0395\u03a3 12\u03a4 \u039d\u0395\u0395\u03a3",
}


class DashboardDataError(RuntimeError):
    pass


class AppStorage:
    def __init__(self, base_directory: Path, storage_dir: Path | None = None):
        self.base_directory = Path(base_directory)
        self.storage_dir = Path(storage_dir) if storage_dir else self.base_directory / DEFAULT_STORAGE_DIRNAME
        self.uploads_dir = self.storage_dir / "uploads"
        self.metadata_path = self.storage_dir / ACTIVE_WORKBOOK_METADATA
        self.storage_dir.mkdir(parents=True, exist_ok=True)
        self.uploads_dir.mkdir(parents=True, exist_ok=True)

    def load_metadata(self) -> dict[str, Any]:
        if not self.metadata_path.exists():
            return {}
        return json.loads(self.metadata_path.read_text(encoding="utf-8"))

    def save_metadata(self, payload: dict[str, Any]) -> None:
        self.metadata_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def resolve_active_workbook(self) -> Path | None:
        metadata = self.load_metadata()
        relative_path = metadata.get("active_workbook")
        if not relative_path:
            return None
        workbook_path = self.storage_dir / relative_path
        if workbook_path.exists():
            return workbook_path
        return None

    def activate_uploaded_workbook(self, filename: str, file_bytes: bytes) -> Path:
        candidate_name = Path(filename).name
        extension = Path(candidate_name).suffix.lower()
        if extension not in ALLOWED_UPLOAD_EXTENSIONS:
            raise DashboardDataError("Επιτρέπονται μόνο αρχεία Excel τύπου .xls, .xlsx ή .xlsm.")
        if not file_bytes:
            raise DashboardDataError("Το αρχείο upload είναι κενό.")
        if len(file_bytes) > MAX_UPLOAD_BYTES:
            raise DashboardDataError("Το αρχείο Excel ξεπερνά το επιτρεπτό όριο μεγέθους.")

        safe_stem = re.sub(r"[^A-Za-z0-9._-]+", "_", Path(candidate_name).stem).strip("._") or "workbook"
        timestamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        final_name = f"{timestamp}-{safe_stem}{extension}"
        final_path = self.uploads_dir / final_name

        with NamedTemporaryFile(delete=False, dir=self.uploads_dir, suffix=extension) as temp_file:
            temp_path = Path(temp_file.name)
            temp_file.write(file_bytes)

        try:
            load_dashboard_payload(temp_path)
        except Exception:
            temp_path.unlink(missing_ok=True)
            raise

        temp_path.replace(final_path)
        self.save_metadata(
            {
                "active_workbook": str(final_path.relative_to(self.storage_dir)),
                "original_filename": candidate_name,
                "uploaded_at": datetime.now().isoformat(timespec="seconds"),
                "size_bytes": len(file_bytes),
            }
        )
        return final_path


def clean_text(value: Any) -> str | None:
    if value is None or pd.isna(value):
        return None
    text = str(value).replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    return text or None


def parse_number(value: Any, default: int = 0) -> int:
    if value is None or pd.isna(value):
        return default
    if isinstance(value, str):
        text = clean_text(value)
        if not text or text == "-":
            return default
        text = text.replace(",", ".")
        try:
            return int(float(text))
        except ValueError:
            return default
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return int(value)
    return default


def format_date(value: Any) -> str | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    text = clean_text(value)
    if not text or text == "00:00:00":
        return None
    return text


def safe_percent(part: int, whole: int) -> float:
    if whole <= 0:
        return 0.0
    return round((part / whole) * 100, 1)


def engine_for_path(path: Path) -> str | None:
    if path.suffix.lower() == ".xls":
        return "xlrd"
    if path.suffix.lower() in {".xlsx", ".xlsm"}:
        return "openpyxl"
    return None


def find_latest_excel_file(directory: Path) -> Path:
    matches = [
        path
        for path in directory.iterdir()
        if path.is_file()
        and path.name.startswith(EXCEL_PREFIX)
        and not path.name.startswith(LOCK_FILE_PREFIX)
        and path.suffix.lower() in {".xls", ".xlsx", ".xlsm"}
    ]
    if not matches:
        raise DashboardDataError(
            f"Δεν βρέθηκε αρχείο Excel που να ξεκινά από '{EXCEL_PREFIX}'."
        )
    return max(matches, key=lambda candidate: (candidate.stat().st_mtime_ns, candidate.name))


def resolve_workbook_path(directory: Path, storage: AppStorage | None = None) -> Path:
    if storage:
        active_workbook = storage.resolve_active_workbook()
        if active_workbook:
            return active_workbook
    return find_latest_excel_file(directory)


def is_authorized_admin(expected_secret: str | None, provided_secret: str | None) -> bool:
    if not expected_secret or not provided_secret:
        return False
    return hmac.compare_digest(expected_secret, provided_secret)


def load_sheet_rows(workbook_path: Path, sheet_name: str | int) -> list[list[Any]]:
    frame = pd.read_excel(
        workbook_path,
        sheet_name=sheet_name,
        header=None,
        engine=engine_for_path(workbook_path),
    )
    return frame.where(pd.notna(frame), None).values.tolist()


def normalize_management_category(raw_name: str, occurrence: int) -> str:
    name = clean_text(raw_name) or ""
    name = MANAGEMENT_CATEGORY_ALIASES.get(name, name)
    if name == "\u03a4\u03a1\u0391\u039a\u03a4\u039f\u03a1\u0395\u03a3":
        return "\u03a4\u03a1\u0391\u039a\u03a4\u039f\u03a1\u0395\u03a3 \u03a0\u0391\u039b\u0399\u039f\u0399" if occurrence == 1 else "\u03a4\u03a1\u0391\u039a\u03a4\u039f\u03a1\u0395\u03a3 LEASING"
    return name


def parse_vehicle_row(category_name: str, row: list[Any]) -> dict[str, Any]:
    vehicle_id = str(parse_number(row[1]))
    available_flag = parse_number(row[2], default=0) == 1
    broken_flag = parse_number(row[3], default=0) == 1
    note = clean_text(row[7]) or "\u0391\u0393\u039d\u03a9\u03a3\u03a4\u039f"
    status = "\u03a3\u03b5 \u03bb\u03b5\u03b9\u03c4\u03bf\u03c5\u03c1\u03b3\u03af\u03b1" if available_flag else "\u039c\u03b5 \u03b2\u03bb\u03ac\u03b2\u03b7" if broken_flag else "\u0386\u03b3\u03bd\u03c9\u03c3\u03c4\u03bf"
    return {
        "row_index": parse_number(row[0]),
        "category_name": category_name,
        "vehicle_id": vehicle_id,
        "available_flag": available_flag,
        "broken_flag": broken_flag,
        "status": status,
        "breakdown_date": format_date(row[4]),
        "expected_return_date": format_date(row[5]),
        "issue_description": clean_text(row[6]),
        "note": note,
        "excluded": False,
        "exclusion_reason": None,
    }


def parse_management_sections(workbook_path: Path) -> dict[str, list[dict[str, Any]]]:
    rows = load_sheet_rows(workbook_path, 2)
    categories: dict[str, list[dict[str, Any]]] = defaultdict(list)
    occurrence_counter: Counter[str] = Counter()
    current_category: str | None = None

    for row in rows:
        marker = clean_text(row[0]) if row else None
        if marker in HEADER_MARKERS:
            raw_name = clean_text(row[3])
            if not raw_name:
                continue
            occurrence_counter[raw_name] += 1
            current_category = normalize_management_category(
                raw_name,
                occurrence_counter[raw_name],
            )
            continue

        if not current_category:
            continue

        if marker == "\u03a3\u03a5\u039d\u039f\u039b\u0391":
            current_category = None
            continue

        if isinstance(row[0], (int, float)) and parse_number(row[1], default=0) > 0:
            categories[current_category].append(parse_vehicle_row(current_category, row))

    if not categories:
        raise DashboardDataError("Δεν εντοπίστηκαν sections οχημάτων στο φύλλο διαχείρισης.")

    return categories


def parse_explicit_exclusions(workbook_path: Path) -> dict[str, str]:
    rows = load_sheet_rows(workbook_path, 3)
    exclusions: dict[str, str] = {}

    for row in rows:
        if isinstance(row[0], (int, float)):
            vehicle_id = str(parse_number(row[1]))
            exclusions[vehicle_id] = clean_text(row[7]) or "Αφαίρεση από την πραγματική απεικόνιση"

    return exclusions


def parse_summary_sheet(workbook_path: Path) -> dict[str, Any]:
    rows = load_sheet_rows(workbook_path, 0)
    if len(rows) < 28:
        raise DashboardDataError("Το φύλλο συνολικής απεικόνισης είναι μικρότερο από το αναμενόμενο.")

    report_date = format_date(rows[1][2])
    raw_time = clean_text(rows[1][7]) or ""
    report_time = raw_time.replace("\u03a9\u03a1\u0391:", "").strip() if raw_time else None

    categories_all: list[dict[str, Any]] = []
    categories_real: list[dict[str, Any]] = []
    real_map: dict[str, dict[str, Any]] = {}

    for row in rows[5:27]:
        total_name = clean_text(row[1])
        if total_name:
            total_count = parse_number(row[4])
            categories_all.append(
                {
                    "name": total_name,
                    "in_service": parse_number(row[2]),
                    "broken": parse_number(row[3]),
                    "total": total_count,
                    "availability_pct": safe_percent(parse_number(row[2]), total_count),
                }
            )

        real_name = clean_text(row[7])
        if real_name:
            real_total = parse_number(row[10])
            payload = {
                "name": real_name,
                "in_service": parse_number(row[8]),
                "broken": parse_number(row[9]),
                "total": real_total,
                "availability_pct": safe_percent(parse_number(row[8]), real_total),
            }
            categories_real.append(payload)
            real_map[real_name] = payload

    if len(categories_all) != 22 or len(categories_real) != 22:
        raise DashboardDataError("Δεν χαρτογραφήθηκαν σωστά και οι 22 κατηγορίες του summary.")

    return {
        "title": clean_text(rows[0][0]) or "Dashboard Κατάστασης Στόλου Οχημάτων",
        "report_date": report_date,
        "report_time": report_time,
        "all_vehicles": {
            "in_service": parse_number(rows[6][13]),
            "broken": parse_number(rows[7][13]),
            "total": parse_number(rows[8][13]),
        },
        "real_fleet": {
            "in_service": parse_number(rows[12][13]),
            "broken": parse_number(rows[13][13]),
            "total": parse_number(rows[14][13]),
        },
        "collection_vehicles_total": parse_number(rows[27][13]),
        "categories_all": categories_all,
        "categories_real": categories_real,
        "categories_real_map": real_map,
    }


def vehicle_sort_key(vehicle: dict[str, Any]) -> tuple[int, int, str, str]:
    note = vehicle["note"]
    note_priority = 0 if note in INFERRED_EXCLUSION_NOTES else 1
    availability_priority = 0 if not vehicle["available_flag"] else 1
    breakdown_date = vehicle["breakdown_date"] or "9999-12-31"
    return (note_priority, availability_priority, breakdown_date, vehicle["vehicle_id"])


def reconcile_vehicle_exclusions(
    categories: dict[str, list[dict[str, Any]]],
    real_summary_map: dict[str, dict[str, Any]],
    explicit_exclusions: dict[str, str],
) -> list[str]:
    warnings: list[str] = []

    for vehicles in categories.values():
        for vehicle in vehicles:
            reason = explicit_exclusions.get(vehicle["vehicle_id"])
            if reason:
                vehicle["excluded"] = True
                vehicle["exclusion_reason"] = reason

    for category_name, vehicles in categories.items():
        summary = real_summary_map.get(category_name)
        if not summary:
            warnings.append(f"Η κατηγορία '{category_name}' υπάρχει στο vehicle sheet αλλά όχι στο summary.")
            continue

        included_now = [vehicle for vehicle in vehicles if not vehicle["excluded"]]
        diff = len(included_now) - summary["total"]
        if diff == 0:
            continue
        if diff < 0:
            warnings.append(
                f"Η κατηγορία '{category_name}' έχει λιγότερα οχήματα στο vehicle sheet ({len(included_now)}) από το summary ({summary['total']})."
            )
            continue

        for vehicle in sorted(included_now, key=vehicle_sort_key)[:diff]:
            vehicle["excluded"] = True
            vehicle["exclusion_reason"] = "Συμπληρωματική αφαίρεση από τη συνοπτική πραγματική απεικόνιση"

    return warnings


def build_category_payload(
    summary_data: dict[str, Any],
    vehicle_map: dict[str, list[dict[str, Any]]],
) -> list[dict[str, Any]]:
    items: list[dict[str, Any]] = []
    all_map = {item["name"]: item for item in summary_data["categories_all"]}

    for order, summary_item in enumerate(summary_data["categories_real"], start=1):
        name = summary_item["name"]
        vehicles = [dict(vehicle) for vehicle in vehicle_map.get(name, [])]
        vehicles.sort(key=lambda vehicle: (vehicle["excluded"], not vehicle["available_flag"], vehicle["vehicle_id"]))

        included = [vehicle for vehicle in vehicles if not vehicle["excluded"]]
        excluded = [vehicle for vehicle in vehicles if vehicle["excluded"]]
        workshop_counts = Counter(
            vehicle["note"]
            for vehicle in included
            if vehicle["note"] and vehicle["note"] != "\u0395\u039d\u0395\u03a1\u0393\u039f"
        )

        items.append(
            {
                "order": order,
                "name": name,
                "summary": summary_item,
                "all_summary": all_map.get(name, summary_item),
                "vehicle_count": len(vehicles),
                "included_vehicle_count": len(included),
                "excluded_vehicle_count": len(excluded),
                "vehicles": vehicles,
                "workshop_counts": [
                    {"name": workshop_name, "count": count}
                    for workshop_name, count in workshop_counts.most_common()
                ],
            }
        )

    return items


def build_alerts(categories: list[dict[str, Any]]) -> dict[str, Any]:
    ranked = sorted(categories, key=lambda item: item["summary"]["availability_pct"], reverse=True)
    best = ranked[:3]
    worst = [item for item in ranked if item["summary"]["availability_pct"] == 0][:3]
    critical = [item for item in ranked if item["summary"]["availability_pct"] < 20]
    return {
        "best_categories": [
            {"name": item["name"], "availability_pct": item["summary"]["availability_pct"]}
            for item in best
        ],
        "worst_categories": [
            {"name": item["name"], "availability_pct": item["summary"]["availability_pct"]}
            for item in worst
        ],
        "critical_count": len(critical),
    }


def build_workshop_summary(categories: list[dict[str, Any]]) -> list[dict[str, Any]]:
    counter: Counter[str] = Counter()
    for category in categories:
        for vehicle in category["vehicles"]:
            note = vehicle["note"]
            if note and note != "\u0395\u039d\u0395\u03a1\u0393\u039f":
                counter[note] += 1
    return [{"name": name, "count": count} for name, count in counter.most_common()]


def load_dashboard_payload(workbook_path: Path) -> dict[str, Any]:
    workbook_path = Path(workbook_path)
    if not workbook_path.exists():
        raise DashboardDataError(f"Το αρχείο '{workbook_path.name}' δεν βρέθηκε.")

    summary_data = parse_summary_sheet(workbook_path)
    vehicle_map = parse_management_sections(workbook_path)
    explicit_exclusions = parse_explicit_exclusions(workbook_path)
    warnings = reconcile_vehicle_exclusions(
        vehicle_map,
        summary_data["categories_real_map"],
        explicit_exclusions,
    )

    categories = build_category_payload(summary_data, vehicle_map)
    workshops = build_workshop_summary(categories)
    meta = {
        "source_file": workbook_path.name,
        "source_path": str(workbook_path),
        "last_modified": datetime.fromtimestamp(workbook_path.stat().st_mtime).isoformat(timespec="seconds"),
        "loaded_at": datetime.now().isoformat(timespec="seconds"),
        "version": f"{workbook_path.stat().st_mtime_ns}:{workbook_path.stat().st_size}",
        "warnings": warnings,
    }

    return {
        "meta": meta,
        "summary": {
            "all_vehicles": summary_data["all_vehicles"],
            "real_fleet": summary_data["real_fleet"],
            "collection_vehicles_total": summary_data["collection_vehicles_total"],
            "report_date": summary_data["report_date"],
            "report_time": summary_data["report_time"],
            "title": summary_data["title"],
        },
        "alerts": build_alerts(categories),
        "categories": categories,
        "workshops": workshops,
    }


def load_latest_dashboard_payload(directory: Path, storage: AppStorage | None = None) -> dict[str, Any]:
    return load_dashboard_payload(resolve_workbook_path(directory, storage))


def build_dashboard_html() -> str:
    return f"""<!DOCTYPE html>
<html lang="el">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Live Dashboard Κατάστασης Στόλου Οχημάτων</title>
  <script src="https://cdn.plot.ly/plotly-2.27.0.min.js"></script>
  <style>
    :root {{
      --bg: #edf2ff;
      --card: #ffffff;
      --ink: #102542;
      --muted: #54657a;
      --line: #d7deeb;
      --accent: #1e3c72;
      --green: #0f9d58;
      --red: #d93025;
      --amber: #f9ab00;
      --blue: #1a73e8;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: "Segoe UI", Tahoma, sans-serif;
      background: linear-gradient(180deg, #dbe7ff 0%, #eef3fb 100%);
      color: var(--ink);
    }}
    .container {{ max-width: 1480px; margin: 0 auto; padding: 24px; }}
    .header {{
      background: linear-gradient(135deg, #143b75 0%, #2f65b9 100%);
      color: white;
      border-radius: 20px;
      padding: 28px;
      box-shadow: 0 20px 40px rgba(20, 59, 117, 0.18);
      margin-bottom: 20px;
    }}
    .header h1 {{ margin: 0 0 10px 0; font-size: 30px; }}
    .header-meta {{ display: flex; flex-wrap: wrap; gap: 18px; font-size: 14px; opacity: 0.95; }}
    .tabs {{ display: flex; gap: 10px; flex-wrap: wrap; margin-bottom: 20px; }}
    .tab {{
      border: 0; background: white; color: var(--accent); padding: 12px 18px;
      border-radius: 999px; font-weight: 700; cursor: pointer;
      box-shadow: 0 6px 16px rgba(25, 63, 119, 0.08);
    }}
    .tab.active {{ background: var(--accent); color: white; }}
    .tab-panel {{ display: none; }}
    .tab-panel.active {{ display: block; }}
    .grid {{ display: grid; gap: 16px; }}
    .grid.kpis {{ grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); margin-bottom: 18px; }}
    .grid.charts {{ grid-template-columns: repeat(auto-fit, minmax(360px, 1fr)); }}
    .card {{
      background: var(--card); border-radius: 18px; padding: 18px;
      box-shadow: 0 10px 24px rgba(16, 37, 66, 0.08);
      border: 1px solid rgba(215, 222, 235, 0.7);
    }}
    .kpi-value {{ font-size: 34px; font-weight: 800; margin-bottom: 6px; }}
    .kpi-label {{ color: var(--muted); font-size: 14px; }}
    .section-title {{ margin: 0 0 14px 0; font-size: 18px; }}
    .note-list {{ margin: 0; padding-left: 18px; color: var(--muted); }}
    .alert {{ border-radius: 14px; padding: 14px 16px; margin-bottom: 14px; font-weight: 600; }}
    .alert.success {{ background: rgba(15, 157, 88, 0.12); color: #0b6b3c; }}
    .alert.warning {{ background: rgba(249, 171, 0, 0.15); color: #8a5f00; }}
    .alert.error {{ background: rgba(217, 48, 37, 0.12); color: #a7271d; }}
    .tree-list {{ display: grid; gap: 12px; }}
    .tree-node {{ border: 1px solid var(--line); border-radius: 14px; background: #fbfdff; overflow: hidden; }}
    .tree-node summary {{
      list-style: none; cursor: pointer; padding: 14px 16px; display: grid;
      grid-template-columns: minmax(220px, 2fr) repeat(5, minmax(90px, 1fr));
      gap: 10px; align-items: center; font-weight: 700;
    }}
    .tree-node summary::-webkit-details-marker {{ display: none; }}
    .tree-node[open] summary {{ background: rgba(26, 115, 232, 0.06); }}
    .summary-cell.muted {{ color: var(--muted); font-weight: 600; }}
    .nested-tree {{ padding: 0 12px 12px 12px; display: grid; gap: 10px; }}
    .group-node {{ border: 1px solid var(--line); border-radius: 12px; overflow: hidden; background: white; }}
    .group-node summary {{ padding: 12px 14px; cursor: pointer; font-weight: 700; background: rgba(20, 59, 117, 0.04); }}
    .vehicle-table {{ width: 100%; border-collapse: collapse; font-size: 14px; }}
    .vehicle-table th, .vehicle-table td {{
      padding: 10px 12px; border-top: 1px solid var(--line); text-align: left; vertical-align: top;
    }}
    .badge {{ display: inline-block; padding: 4px 10px; border-radius: 999px; font-size: 12px; font-weight: 700; }}
    .badge.good {{ background: rgba(15, 157, 88, 0.14); color: #0b6b3c; }}
    .badge.warn {{ background: rgba(249, 171, 0, 0.16); color: #8a5f00; }}
    .badge.bad {{ background: rgba(217, 48, 37, 0.14); color: #a7271d; }}
    .meta-box {{ font-size: 14px; color: var(--muted); line-height: 1.6; }}
    .status-line {{ margin-bottom: 8px; color: var(--muted); font-size: 14px; }}
    .chart {{ min-height: 320px; }}
    .empty {{ padding: 14px; color: var(--muted); font-style: italic; }}
    @media (max-width: 980px) {{
      .tree-node summary {{ grid-template-columns: 1fr 1fr; }}
    }}
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1 id="dashboardTitle">Live Dashboard Κατάστασης Στόλου Οχημάτων</h1>
      <div class="header-meta">
        <span id="reportDate">Ημερομηνία: -</span>
        <span id="reportTime">Ώρα: -</span>
        <span id="sourceFile">Αρχείο: -</span>
        <span id="lastLoaded">Τελευταία φόρτωση: -</span>
      </div>
    </div>

    <div class="tabs">
      <button class="tab active" data-tab="overview">Επισκόπηση</button>
      <button class="tab" data-tab="availability">Διαθεσιμότητα</button>
      <button class="tab" data-tab="workshops">Συνεργεία</button>
      <button class="tab" data-tab="details">Λεπτομέρειες</button>
      <button class="tab" data-tab="source">Πηγή και Έλεγχοι</button>
    </div>

    <div id="globalMessage"></div>

    <section class="tab-panel active" id="overview">
      <div class="grid kpis" id="overviewKpis"></div>
      <div class="grid charts">
        <div class="card">
          <h2 class="section-title">Κατάσταση πραγματικού στόλου</h2>
          <div id="statusChart" class="chart"></div>
        </div>
        <div class="card">
          <h2 class="section-title">Top διαθεσιμότητα ανά κατηγορία</h2>
          <div id="availabilityChart" class="chart"></div>
        </div>
      </div>
    </section>

    <section class="tab-panel" id="availability">
      <div id="availabilityAlerts"></div>
      <div class="grid charts">
        <div class="card">
          <h2 class="section-title">Διαθεσιμότητα πραγματικού στόλου</h2>
          <div id="gaugeChart" class="chart"></div>
        </div>
        <div class="card">
          <h2 class="section-title">Κατηγορίες προς διερεύνηση</h2>
          <div id="criticalChart" class="chart"></div>
        </div>
      </div>
      <div class="card" style="margin-top: 18px;">
        <h2 class="section-title">Δενδροειδής προβολή διαθεσιμότητας</h2>
        <div class="status-line">Κάνε κλικ σε κάθε κατηγορία για να ανοίξουν ομάδες οχημάτων και αναλυτικές γραμμές.</div>
        <div id="availabilityTree" class="tree-list"></div>
      </div>
    </section>

    <section class="tab-panel" id="workshops">
      <div class="grid charts">
        <div class="card">
          <h2 class="section-title">Κατανομή οχημάτων εκτός λειτουργίας</h2>
          <div id="workshopChart" class="chart"></div>
        </div>
        <div class="card">
          <h2 class="section-title">Πλήθος εκτός λειτουργίας ανά κατηγορία</h2>
          <div id="brokenByCategoryChart" class="chart"></div>
        </div>
      </div>
      <div class="card" style="margin-top: 18px;">
        <h2 class="section-title">Σύνοψη συνεργείων / κατάστασης</h2>
        <div id="workshopSummary"></div>
      </div>
    </section>

    <section class="tab-panel" id="details">
      <div class="card">
        <h2 class="section-title">Δενδροειδής προβολή λεπτομερειών</h2>
        <div class="status-line">Κάθε κατηγορία ανοίγει σε ομάδες και αναλυτικές γραμμές οχημάτων με βλάβη, ημερομηνίες και παρατηρήσεις.</div>
        <div id="detailsTree" class="tree-list"></div>
      </div>
    </section>

    <section class="tab-panel" id="source">
      <div class="card meta-box">
        <h2 class="section-title">Πηγή και έλεγχοι</h2>
        <div id="sourceMeta"></div>
        <div style="margin: 12px 0;">
          <a href="/admin">Admin upload Excel</a>
        </div>
        <div id="warningList"></div>
      </div>
    </section>
  </div>

  <script>
    let lastVersion = null;

    document.querySelectorAll('.tab').forEach((button) => {{
      button.addEventListener('click', () => {{
        document.querySelectorAll('.tab').forEach((item) => item.classList.remove('active'));
        document.querySelectorAll('.tab-panel').forEach((panel) => panel.classList.remove('active'));
        button.classList.add('active');
        document.getElementById(button.dataset.tab).classList.add('active');
      }});
    }});

    function formatPercent(value) {{
      return `${{Number(value || 0).toFixed(1)}}%`;
    }}

    function availabilityBadge(value) {{
      if (value >= 50) return '<span class="badge good">Καλή</span>';
      if (value >= 20) return '<span class="badge warn">Μέτρια</span>';
      return '<span class="badge bad">Κρίσιμη</span>';
    }}

    function escapeHtml(value) {{
      return String(value ?? '')
        .replaceAll('&', '&amp;')
        .replaceAll('<', '&lt;')
        .replaceAll('>', '&gt;')
        .replaceAll('"', '&quot;')
        .replaceAll("'", '&#39;');
    }}

    function groupVehicles(category) {{
      return [
        {{ key: 'includedActive', title: 'Σε λειτουργία', items: category.vehicles.filter((item) => !item.excluded && item.available_flag) }},
        {{ key: 'includedBroken', title: 'Με βλάβη', items: category.vehicles.filter((item) => !item.excluded && !item.available_flag) }},
        {{ key: 'excluded', title: 'Εξαιρέθηκαν από την πραγματική απεικόνιση', items: category.vehicles.filter((item) => item.excluded) }},
      ].filter((group) => group.items.length > 0);
    }}

    function renderVehicleRows(vehicles) {{
      if (!vehicles.length) {{
        return '<div class="empty">Δεν υπάρχουν αναλυτικές γραμμές.</div>';
      }}
      const rows = vehicles.map((vehicle) => `
        <tr>
          <td>${{escapeHtml(vehicle.vehicle_id)}}</td>
          <td>${{escapeHtml(vehicle.status)}}</td>
          <td>${{escapeHtml(vehicle.note)}}</td>
          <td>${{escapeHtml(vehicle.issue_description || '')}}</td>
          <td>${{escapeHtml(vehicle.breakdown_date || '')}}</td>
          <td>${{escapeHtml(vehicle.expected_return_date || '')}}</td>
          <td>${{escapeHtml(vehicle.exclusion_reason || '')}}</td>
        </tr>
      `).join('');
      return `
        <table class="vehicle-table">
          <thead>
            <tr>
              <th>Κωδικός</th>
              <th>Κατάσταση</th>
              <th>Παρατήρηση</th>
              <th>Βλάβη</th>
              <th>Ημ. βλάβης</th>
              <th>Εκτ. αποκατάσταση</th>
              <th>Σημείωση εξαίρεσης</th>
            </tr>
          </thead>
          <tbody>${{rows}}</tbody>
        </table>
      `;
    }}

    function renderTree(containerId, categories) {{
      const container = document.getElementById(containerId);
      if (!categories.length) {{
        container.innerHTML = '<div class="empty">Δεν βρέθηκαν κατηγορίες.</div>';
        return;
      }}

      container.innerHTML = categories.map((category) => {{
        const groups = groupVehicles(category);
        const nested = groups.map((group) => `
          <details class="group-node">
            <summary>${{escapeHtml(group.title)}} (${{group.items.length}})</summary>
            ${{renderVehicleRows(group.items)}}
          </details>
        `).join('');

        return `
          <details class="tree-node">
            <summary>
              <span>${{escapeHtml(category.name)}}</span>
              <span class="summary-cell">${{category.summary.total}}</span>
              <span class="summary-cell">${{category.summary.in_service}}</span>
              <span class="summary-cell">${{category.summary.broken}}</span>
              <span class="summary-cell">${{formatPercent(category.summary.availability_pct)}}</span>
              <span class="summary-cell muted">${{availabilityBadge(category.summary.availability_pct)}}</span>
            </summary>
            <div class="nested-tree">${{nested || '<div class="empty">Δεν υπάρχουν αναλυτικά δεδομένα οχημάτων.</div>'}}</div>
          </details>
        `;
      }}).join('');
    }}

    function renderOverview(data) {{
      const real = data.summary.real_fleet;
      const all = data.summary.all_vehicles;
      const kpis = [
        ['Σύνολο πραγματικού στόλου', real.total],
        ['Σε λειτουργία', `${{real.in_service}} (${{formatPercent(real.in_service / real.total * 100)}})`],
        ['Με βλάβη', `${{real.broken}} (${{formatPercent(real.broken / real.total * 100)}})`],
        ['Σύνολο οχημάτων φύλλου', all.total],
        ['Ημερήσια αποκομιδή', data.summary.collection_vehicles_total],
      ];

      document.getElementById('overviewKpis').innerHTML = kpis.map(([label, value]) => `
        <div class="card">
          <div class="kpi-value">${{escapeHtml(value)}}</div>
          <div class="kpi-label">${{escapeHtml(label)}}</div>
        </div>
      `).join('');

      Plotly.react('statusChart', [{{
        values: [real.in_service, real.broken],
        labels: ['Σε λειτουργία', 'Με βλάβη'],
        type: 'pie',
        marker: {{ colors: ['#0f9d58', '#d93025'] }},
        textinfo: 'label+percent'
      }}], {{ margin: {{ t: 20, b: 20, l: 20, r: 20 }} }}, {{ responsive: true }});

      const top = [...data.categories]
        .sort((left, right) => right.summary.availability_pct - left.summary.availability_pct)
        .slice(0, 10);

      Plotly.react('availabilityChart', [{{
        type: 'bar',
        x: top.map((item) => item.name),
        y: top.map((item) => item.summary.availability_pct),
        marker: {{
          color: top.map((item) => item.summary.availability_pct >= 50 ? '#0f9d58' : item.summary.availability_pct >= 20 ? '#f9ab00' : '#d93025')
        }}
      }}], {{
        margin: {{ t: 20, b: 120, l: 50, r: 20 }},
        xaxis: {{ tickangle: -35 }},
        yaxis: {{ title: 'Διαθεσιμότητα %', range: [0, 100] }}
      }}, {{ responsive: true }});
    }}

    function renderAvailability(data) {{
      const real = data.summary.real_fleet;
      const alerts = [];
      if (data.alerts.best_categories.length) {{
        alerts.push(`<div class="alert success">Καλύτερες κατηγορίες: ${{data.alerts.best_categories.map((item) => `${{escapeHtml(item.name)}} (${{formatPercent(item.availability_pct)}})`).join(', ')}}</div>`);
      }}
      if (data.alerts.worst_categories.length) {{
        alerts.push(`<div class="alert error">Μηδενική διαθεσιμότητα: ${{data.alerts.worst_categories.map((item) => escapeHtml(item.name)).join(', ')}}</div>`);
      }}
      alerts.push(`<div class="alert warning">Κατηγορίες κάτω από 20% διαθεσιμότητα: ${{data.alerts.critical_count}}</div>`);
      document.getElementById('availabilityAlerts').innerHTML = alerts.join('');

      Plotly.react('gaugeChart', [{{
        type: 'indicator',
        mode: 'gauge+number',
        value: Number((real.in_service / real.total * 100).toFixed(1)),
        number: {{ suffix: '%' }},
        gauge: {{
          axis: {{ range: [0, 100] }},
          bar: {{ color: '#1a73e8' }},
          steps: [
            {{ range: [0, 20], color: '#f8d7da' }},
            {{ range: [20, 50], color: '#fff3cd' }},
            {{ range: [50, 100], color: '#d4edda' }}
          ]
        }}
      }}], {{ margin: {{ t: 30, b: 30, l: 20, r: 20 }} }}, {{ responsive: true }});

      const critical = [...data.categories]
        .sort((left, right) => left.summary.availability_pct - right.summary.availability_pct)
        .slice(0, 10);

      Plotly.react('criticalChart', [{{
        type: 'bar',
        orientation: 'h',
        x: critical.map((item) => item.summary.availability_pct).reverse(),
        y: critical.map((item) => item.name).reverse(),
        marker: {{ color: '#d93025' }}
      }}], {{
        margin: {{ t: 20, b: 20, l: 180, r: 30 }},
        xaxis: {{ title: 'Διαθεσιμότητα %', range: [0, 100] }}
      }}, {{ responsive: true }});

      renderTree('availabilityTree', data.categories);
    }}

    function renderWorkshops(data) {{
      const workshops = data.workshops;
      Plotly.react('workshopChart', [{{
        type: 'bar',
        x: workshops.map((item) => item.name),
        y: workshops.map((item) => item.count),
        marker: {{ color: '#1a73e8' }}
      }}], {{
        margin: {{ t: 20, b: 100, l: 50, r: 20 }},
        xaxis: {{ tickangle: -30 }}
      }}, {{ responsive: true }});

      Plotly.react('brokenByCategoryChart', [{{
        type: 'bar',
        x: data.categories.map((item) => item.name),
        y: data.categories.map((item) => item.summary.broken),
        marker: {{ color: '#f9ab00' }}
      }}], {{
        margin: {{ t: 20, b: 120, l: 50, r: 20 }},
        xaxis: {{ tickangle: -35 }}
      }}, {{ responsive: true }});

      document.getElementById('workshopSummary').innerHTML = workshops.length
        ? renderVehicleRows(workshops.map((item) => ({{
            vehicle_id: item.count,
            status: 'Εκτός λειτουργίας',
            note: item.name,
            issue_description: '',
            breakdown_date: '',
            expected_return_date: '',
            exclusion_reason: ''
          }})))
        : '<div class="empty">Δεν βρέθηκαν εγγραφές συνεργείων.</div>';
    }}

    function renderDetails(data) {{
      renderTree('detailsTree', data.categories);
    }}

    function renderMeta(data) {{
      document.getElementById('dashboardTitle').textContent = data.summary.title || 'Live Dashboard Κατάστασης Στόλου Οχημάτων';
      document.getElementById('reportDate').textContent = `Ημερομηνία: ${{data.summary.report_date || '-'}}`;
      document.getElementById('reportTime').textContent = `Ώρα: ${{data.summary.report_time || '-'}}`;
      document.getElementById('sourceFile').textContent = `Αρχείο: ${{data.meta.source_file}}`;
      document.getElementById('lastLoaded').textContent = `Τελευταία φόρτωση: ${{data.meta.loaded_at}}`;

      document.getElementById('sourceMeta').innerHTML = `
        <div><strong>Αρχείο:</strong> ${{escapeHtml(data.meta.source_file)}}</div>
        <div><strong>Διαδρομή:</strong> ${{escapeHtml(data.meta.source_path)}}</div>
        <div><strong>Τροποποίηση αρχείου:</strong> ${{escapeHtml(data.meta.last_modified)}}</div>
        <div><strong>Φόρτωση dashboard:</strong> ${{escapeHtml(data.meta.loaded_at)}}</div>
      `;

      const warningList = document.getElementById('warningList');
      if (data.meta.warnings.length) {{
        warningList.innerHTML = `
          <div class="alert warning">
            Ο parser βρήκε προειδοποιήσεις:
            <ul class="note-list">${{data.meta.warnings.map((item) => `<li>${{escapeHtml(item)}}</li>`).join('')}}</ul>
          </div>
        `;
      }} else {{
        warningList.innerHTML = '<div class="alert success">Δεν υπάρχουν προειδοποιήσεις parsing.</div>';
      }}
    }}

    function renderData(data) {{
      renderMeta(data);
      renderOverview(data);
      renderAvailability(data);
      renderWorkshops(data);
      renderDetails(data);
      document.getElementById('globalMessage').innerHTML = '';
    }}

    function renderError(errorPayload) {{
      const message = errorPayload?.error?.message || 'Αδυναμία φόρτωσης του dashboard.';
      const details = errorPayload?.error?.details || [];
      document.getElementById('globalMessage').innerHTML = `
        <div class="alert error">
          <strong>Σφάλμα δεδομένων:</strong> ${{escapeHtml(message)}}
          ${{details.length ? `<ul class="note-list">${{details.map((item) => `<li>${{escapeHtml(item)}}</li>`).join('')}}</ul>` : ''}}
        </div>
      `;
    }}

    async function loadDashboardData(force = false) {{
      try {{
        const response = await fetch(`/api/fleet-data?ts=${{Date.now()}}`);
        const payload = await response.json();
        if (!response.ok) {{
          renderError(payload);
          return;
        }}
        if (!force && lastVersion === payload.meta.version) {{
          return;
        }}
        lastVersion = payload.meta.version;
        renderData(payload);
      }} catch (error) {{
        renderError({{ error: {{ message: error.message, details: [] }} }});
      }}
    }}

    loadDashboardData(true);
    setInterval(loadDashboardData, {POLL_INTERVAL_MS});
  </script>
</body>
</html>
"""


def build_admin_html() -> str:
    return """<!DOCTYPE html>
<html lang="el">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Admin Upload Excel</title>
  <style>
    body {
      font-family: "Segoe UI", Tahoma, sans-serif;
      margin: 0;
      background: #eef3fb;
      color: #102542;
    }
    .wrap {
      max-width: 760px;
      margin: 48px auto;
      background: white;
      border-radius: 18px;
      padding: 28px;
      box-shadow: 0 12px 28px rgba(16, 37, 66, 0.1);
    }
    h1 { margin-top: 0; }
    label { display: block; margin: 14px 0 8px; font-weight: 700; }
    input[type="password"], input[type="file"] {
      width: 100%;
      padding: 12px;
      border: 1px solid #ccd6e2;
      border-radius: 10px;
      box-sizing: border-box;
      background: white;
    }
    button {
      margin-top: 16px;
      border: 0;
      background: #1e3c72;
      color: white;
      padding: 12px 18px;
      border-radius: 10px;
      cursor: pointer;
      font-weight: 700;
    }
    .status {
      margin-top: 18px;
      padding: 12px 14px;
      border-radius: 10px;
      white-space: pre-wrap;
    }
    .status.info { background: #e8f0fe; color: #174ea6; }
    .status.error { background: #fdecea; color: #b3261e; }
    .status.success { background: #e6f4ea; color: #137333; }
    .helper {
      color: #54657a;
      margin-top: 10px;
      line-height: 1.5;
    }
  </style>
</head>
<body>
  <div class="wrap">
    <h1>Admin upload Excel</h1>
    <p class="helper">Ανέβασε νέο Excel για να γίνει το ενεργό workbook του online dashboard. Αν το parse αποτύχει, το τρέχον ενεργό αρχείο δεν αλλάζει.</p>
    <form id="uploadForm">
      <label for="secret">Admin secret</label>
      <input id="secret" type="password" autocomplete="current-password" />

      <label for="workbook">Excel αρχείο</label>
      <input id="workbook" type="file" accept=".xls,.xlsx,.xlsm" />

      <button type="submit">Ανέβασμα νέου Excel</button>
    </form>
    <div id="statusBox" class="status info">Περιμένω νέο αρχείο.</div>
    <p class="helper"><a href="/">Επιστροφή στο dashboard</a></p>
  </div>

  <script>
    const form = document.getElementById('uploadForm');
    const secretInput = document.getElementById('secret');
    const workbookInput = document.getElementById('workbook');
    const statusBox = document.getElementById('statusBox');

    function setStatus(message, type) {
      statusBox.className = `status ${type}`;
      statusBox.textContent = message;
    }

    form.addEventListener('submit', async (event) => {
      event.preventDefault();
      const file = workbookInput.files[0];
      const secret = secretInput.value.trim();

      if (!secret) {
        setStatus('Χρειάζεται admin secret.', 'error');
        return;
      }
      if (!file) {
        setStatus('Διάλεξε ένα Excel αρχείο πριν το upload.', 'error');
        return;
      }

      setStatus('Ανεβάζω και ελέγχω το νέο Excel...', 'info');

      try {
        const response = await fetch('/api/admin/upload', {
          method: 'POST',
          headers: {
            'Content-Type': 'application/octet-stream',
            'X-Admin-Secret': secret,
            'X-Upload-Filename': encodeURIComponent(file.name)
          },
          body: await file.arrayBuffer()
        });

        const payload = await response.json();
        if (!response.ok) {
          setStatus(payload.error?.message || 'Αποτυχία upload.', 'error');
          return;
        }

        const lines = [
          'Το νέο workbook ενεργοποιήθηκε επιτυχώς.',
          `Αρχικό όνομα: ${payload.upload.original_filename}`,
          `Ενεργό αρχείο: ${payload.upload.active_file}`,
          `Φόρτωση: ${payload.upload.uploaded_at}`
        ];
        setStatus(lines.join('\\n'), 'success');
      } catch (error) {
        setStatus(`Σφάλμα upload: ${error.message}`, 'error');
      }
    });
  </script>
</body>
</html>
"""


class DashboardRequestHandler(BaseHTTPRequestHandler):
    base_directory = Path(__file__).resolve().parent
    storage: AppStorage | None = None
    admin_secret = os.getenv(DEFAULT_ADMIN_SECRET_ENV, "")

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/":
            self._write_html(build_dashboard_html())
            return
        if parsed.path == "/admin":
            self._write_html(build_admin_html())
            return
        if parsed.path == "/api/fleet-data":
            self._write_payload()
            return
        self.send_error(HTTPStatus.NOT_FOUND, "Η διαδρομή δεν βρέθηκε.")

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path == "/api/admin/upload":
            self._handle_admin_upload()
            return
        self.send_error(HTTPStatus.NOT_FOUND, "Η διαδρομή δεν βρέθηκε.")

    def log_message(self, format_string: str, *args: Any) -> None:
        return

    def _storage_instance(self) -> AppStorage:
        if self.storage is None:
            type(self).storage = AppStorage(self.base_directory)
        return type(self).storage

    def _write_html(self, html: str) -> None:
        content = html.encode("utf-8")
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def _write_json(self, payload: dict[str, Any], status: HTTPStatus) -> None:
        content = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Cache-Control", "no-store")
        self.send_header("Content-Length", str(len(content)))
        self.end_headers()
        self.wfile.write(content)

    def _write_payload(self) -> None:
        try:
            payload = load_latest_dashboard_payload(self.base_directory, self._storage_instance())
            self._write_json(payload, HTTPStatus.OK)
        except Exception as exc:  # pragma: no cover
            self._write_json(
                {
                    "error": {
                        "message": "Αδυναμία ανάγνωσης του Excel για το dashboard.",
                        "details": [str(exc)],
                    }
                },
                HTTPStatus.SERVICE_UNAVAILABLE,
            )

    def _handle_admin_upload(self) -> None:
        if not is_authorized_admin(self.admin_secret, self.headers.get("X-Admin-Secret")):
            self._write_json(
                {"error": {"message": "Μη εξουσιοδοτημένο upload. Το admin secret δεν είναι έγκυρο."}},
                HTTPStatus.UNAUTHORIZED,
            )
            return

        content_length = parse_number(self.headers.get("Content-Length"), default=0)
        if content_length <= 0:
            self._write_json(
                {"error": {"message": "Δεν ελήφθη περιεχόμενο αρχείου για upload."}},
                HTTPStatus.BAD_REQUEST,
            )
            return
        if content_length > MAX_UPLOAD_BYTES:
            self._write_json(
                {"error": {"message": "Το αρχείο Excel ξεπερνά το επιτρεπτό όριο μεγέθους."}},
                HTTPStatus.REQUEST_ENTITY_TOO_LARGE,
            )
            return

        filename = unquote(self.headers.get("X-Upload-Filename") or "upload.xlsx")
        file_bytes = self.rfile.read(content_length)
        try:
            storage = self._storage_instance()
            workbook_path = storage.activate_uploaded_workbook(filename, file_bytes)
            metadata = storage.load_metadata()
            self._write_json(
                {
                    "ok": True,
                    "upload": {
                        "original_filename": metadata.get("original_filename", filename),
                        "active_file": workbook_path.name,
                        "uploaded_at": metadata.get("uploaded_at"),
                        "size_bytes": metadata.get("size_bytes"),
                    },
                },
                HTTPStatus.CREATED,
            )
        except Exception as exc:
            self._write_json(
                {"error": {"message": str(exc)}},
                HTTPStatus.BAD_REQUEST,
            )


def build_server(
    host: str = DEFAULT_HOST,
    port: int = DEFAULT_PORT,
    base_directory: Path | None = None,
    storage_dir: Path | None = None,
    admin_secret: str | None = None,
) -> ThreadingHTTPServer:
    DashboardRequestHandler.base_directory = base_directory or Path(__file__).resolve().parent
    DashboardRequestHandler.storage = AppStorage(DashboardRequestHandler.base_directory, storage_dir)
    DashboardRequestHandler.admin_secret = admin_secret if admin_secret is not None else os.getenv(DEFAULT_ADMIN_SECRET_ENV, "")
    return ThreadingHTTPServer((host, port), DashboardRequestHandler)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Live dashboard για το Excel κατάστασης στόλου οχημάτων.")
    parser.add_argument("--host", default=DEFAULT_HOST)
    parser.add_argument("--port", type=int, default=DEFAULT_PORT)
    parser.add_argument(
        "--once",
        action="store_true",
        help="Εκτύπωσε μόνο το JSON payload του ενεργού Excel και τερμάτισε.",
    )
    parser.add_argument(
        "--storage-dir",
        default=None,
        help="Directory για active workbook metadata και uploaded Excel αρχεία.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    base_directory = Path(__file__).resolve().parent
    storage_dir = Path(args.storage_dir) if args.storage_dir else None
    storage = AppStorage(base_directory, storage_dir)
    if args.once:
        payload = load_latest_dashboard_payload(base_directory, storage)
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return

    server = build_server(
        args.host,
        args.port,
        base_directory=base_directory,
        storage_dir=storage_dir,
    )
    url = f"http://{args.host}:{args.port}"
    print(f"Dashboard διαθέσιμο στο {url}")
    print("Δημόσιο dashboard: /")
    print("Admin upload UI: /admin")
    print(
        f"Αν υπάρχει ενεργό uploaded workbook θα χρησιμοποιηθεί, αλλιώς θα γίνει fallback στο νεότερο Excel με πρόθεμα '{EXCEL_PREFIX}'."
    )
    server.serve_forever()


if __name__ == "__main__":
    main()

