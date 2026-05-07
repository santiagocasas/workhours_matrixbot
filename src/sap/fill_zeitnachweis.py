from __future__ import annotations

import argparse
import calendar
import json
from dataclasses import dataclass
from datetime import date, datetime
from pathlib import Path
from typing import Any

import yaml
from openpyxl import load_workbook
from playwright.sync_api import Page, sync_playwright

from src.excel.handler import BRIDGE_DAY_CODE, HOLIDAY_CODE, MONTH_SHEETS
from src.excel.time_utils import compute_preview_hours, normalize_excel_time


SAP_URL = "https://webpostkorb.dlr.de/Antrag/NeuerAntrag?gvTypKuerzel=ZERF&gvSubTyp=&gvTypName=Zeitnachweis&hatProfil=true"
ABSENCE_CODES = {"U", "K"}
SKIP_CODES = {HOLIDAY_CODE, BRIDGE_DAY_CODE}


@dataclass(frozen=True)
class SapDayEntry:
    day: int
    code: str | None
    value: str
    target: str


def load_config() -> dict[str, Any]:
    config_path = Path(__file__).resolve().parents[2] / "config.yaml"
    with config_path.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle)


def read_month_entries(workbook_path: str, year: int, month: int) -> list[SapDayEntry]:
    workbook = load_workbook(workbook_path, data_only=True, read_only=True)
    try:
        sheet = workbook[MONTH_SHEETS[month]]
        holiday_dates = _collect_date_set(workbook["Allgemein"], "L", 4, 27)
        bridge_dates = _collect_date_set(workbook["Allgemein"], "L", 32, 40)
        entries: list[SapDayEntry] = []

        for row in range(8, 46):
            cell_date = sheet[f"A{row}"].value
            if not _matches_month_date(cell_date, year, month):
                continue

            target_date = cell_date.date() if isinstance(cell_date, datetime) else cell_date
            code = _effective_code(sheet[f"S{row}"].value, target_date, holiday_dates, bridge_dates)
            value = _format_sap_number(sheet[f"T{row}"].value)
            if not value:
                value = _computed_sap_value(workbook, sheet, row, target_date, code)
            day = target_date.day

            if not value:
                continue
            if code in SKIP_CODES:
                continue
            target = "absence" if code in ABSENCE_CODES else "project"
            entries.append(SapDayEntry(day=day, code=code, value=value, target=target))

        return entries
    finally:
        workbook.close()


def _matches_month_date(value: object, year: int, month: int) -> bool:
    if isinstance(value, datetime):
        return value.year == year and value.month == month
    if isinstance(value, date):
        return value.year == year and value.month == month
    return False


def _normalize_code(value: object) -> str | None:
    if not isinstance(value, str):
        return None
    cleaned = value.strip().upper()
    if not cleaned or cleaned.startswith("="):
        return None
    return cleaned


def _effective_code(
    value: object,
    target_date: date,
    holiday_dates: set[date],
    bridge_dates: set[date],
) -> str | None:
    code = _normalize_code(value)
    if code:
        return code
    if target_date in holiday_dates:
        return HOLIDAY_CODE
    if target_date in bridge_dates:
        return BRIDGE_DAY_CODE
    return None


def _collect_date_set(sheet, column: str, start_row: int, end_row: int) -> set[date]:
    values: set[date] = set()
    for row in range(start_row, end_row + 1):
        value = sheet[f"{column}{row}"].value
        if isinstance(value, datetime):
            values.add(value.date())
        elif isinstance(value, date):
            values.add(value)
    return values


def _computed_sap_value(workbook, sheet, row: int, target_date: date, code: str | None) -> str:
    start = normalize_excel_time(sheet[f"C{row}"].value)
    end = normalize_excel_time(sheet[f"D{row}"].value)
    breaks = []
    for start_col, end_col in (("F", "G"), ("H", "I"), ("J", "K")):
        break_start = normalize_excel_time(sheet[f"{start_col}{row}"].value)
        break_end = normalize_excel_time(sheet[f"{end_col}{row}"].value)
        if break_start is not None and break_end is not None:
            breaks.append((break_start, break_end))

    soll = normalize_excel_time(workbook["Allgemein"][f"C{26 + target_date.weekday()}"].value)
    if soll is None:
        return ""
    value = compute_preview_hours(start, end, breaks, code, soll)
    return _format_sap_number(value)


def _format_sap_number(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, str):
        cleaned = value.strip()
        return cleaned.replace(".", ",")
    if isinstance(value, (int, float)):
        if abs(float(value)) < 0.0001:
            return ""
        return f"{float(value):.1f}".replace(".", ",")
    return ""


def fill_sap_form(
    page: Page,
    entries: list[SapDayEntry],
    cost_center: str,
    project_row: int,
) -> None:
    page.fill(f"#ZN_KTR\\.{project_row}\\.KTR_Nr", cost_center)
    page.keyboard.press("Tab")

    for entry in entries:
        selector = _selector_for_entry(entry, project_row)
        page.fill(selector, entry.value)
        page.keyboard.press("Tab")


def _selector_for_entry(entry: SapDayEntry, project_row: int) -> str:
    day = f"{entry.day:02d}"
    if entry.target == "absence":
        return f"#AB{day}"
    return f"#ZN_KTR\\.{project_row}\\.KTR{day}"


def print_plan(entries: list[SapDayEntry], cost_center: str, year: int, month: int) -> None:
    print(f"Ktr-Nr.: {cost_center}")
    if not entries:
        print("No SAP entries found for this month.")
        return
    print_horizontal_plan(entries, year, month)
    print()
    print("Detailed values:")
    for entry in entries:
        label = "AB" if entry.target == "absence" else "KTR"
        code = f" code={entry.code}" if entry.code else ""
        print(f"day {entry.day:02d}: {label} {entry.value}{code}")


def print_horizontal_plan(entries: list[SapDayEntry], year: int, month: int) -> None:
    days_in_month = calendar.monthrange(year, month)[1]
    absence = _values_by_day(entries, "absence")
    project = _values_by_day(entries, "project")
    label_width = 13

    print()
    print("SAP table view:")
    print(_horizontal_row("Day", [f"{day:02d}" for day in range(1, days_in_month + 1)], label_width))
    print(_horizontal_row("Urlaub/Krank", [absence.get(day, "") for day in range(1, days_in_month + 1)], label_width))
    print(_horizontal_row("Section II", [project.get(day, "") for day in range(1, days_in_month + 1)], label_width))


def _values_by_day(entries: list[SapDayEntry], target: str) -> dict[int, str]:
    return {entry.day: entry.value for entry in entries if entry.target == target}


def _horizontal_row(label: str, values: list[str], label_width: int) -> str:
    return f"{label:<{label_width}}" + " ".join(f"{value:>4}" for value in values)


def print_browser_console_instructions(copied: bool) -> None:
    source = "The browser JavaScript is already copied to the Windows clipboard." if copied else "Copy the printed browser JavaScript first."
    print()
    print("Windows Firefox instructions:")
    print(f"1. {source}")
    print("2. Open the SAP Zeitnachweis form in your normal Windows Firefox.")
    print("3. Check that the correct month is selected.")
    print("4. Press F12.")
    print("5. Open the Console tab.")
    print("6. If Firefox blocks pasting, type: allow pasting")
    print("7. Paste with Ctrl+V and press Enter.")
    print("8. Review all filled fields before saving manually.")


def browser_script(entries: list[SapDayEntry], cost_center: str, project_row: int) -> str:
    payload = [
        {
            "id": _dom_id_for_entry(entry, project_row),
            "value": entry.value,
            "label": f"day {entry.day:02d} {entry.target}",
        }
        for entry in entries
    ]
    payload.insert(
        0,
        {
            "id": f"ZN_KTR.{project_row}.KTR_Nr",
            "value": cost_center,
            "label": "Ktr-Nr.",
        },
    )
    return """(() => {
  const fields = __FIELDS__;
  const missing = [];
  const changed = [];

  function documentsToSearch() {
    const docs = [document];
    for (let index = 0; index < window.frames.length; index += 1) {
      const frame = window.frames[index];
      try {
        if (frame.document) {
          docs.push(frame.document);
        }
      } catch (error) {
        console.warn('Zeitnachweis cannot access frame:', error);
      }
    }
    return docs;
  }

  function findField(id) {
    for (const doc of documentsToSearch()) {
      const el = doc.getElementById(id);
      if (el) {
        return el;
      }
    }
    return null;
  }

  function setField(id, value) {
    const el = findField(id);
    if (!el) {
      missing.push(id);
      return;
    }
    el.focus();
    el.value = value;
    el.dispatchEvent(new Event('input', { bubbles: true }));
    el.dispatchEvent(new Event('change', { bubbles: true }));
    el.dispatchEvent(new Event('blur', { bubbles: true }));
    changed.push(`${id}=${value}`);
  }

  for (const field of fields) {
    setField(field.id, field.value);
  }

  console.log('Zeitnachweis changed fields:', changed);
  console.log('Zeitnachweis missing fields:', missing);
  if (missing.length) {
    alert(`Zeitnachweis filled ${changed.length} fields, but ${missing.length} fields were not found. See console for details.`);
  } else {
    alert(`Zeitnachweis filled ${changed.length} fields. Please review before saving.`);
  }
})();""".replace("__FIELDS__", json.dumps(payload, ensure_ascii=False, indent=2))


def _dom_id_for_entry(entry: SapDayEntry, project_row: int) -> str:
    day = f"{entry.day:02d}"
    if entry.target == "absence":
        return f"AB{day}"
    return f"ZN_KTR.{project_row}.KTR{day}"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Fill DLR SAP Zeitnachweis from Excel.")
    parser.add_argument("--year", type=int, default=datetime.now().year)
    parser.add_argument("--month", type=int, required=True)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--browser-js", action="store_true")
    parser.add_argument("--copy", action="store_true")
    parser.add_argument("--url")
    parser.add_argument("--profile-path")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    config = load_config()
    sap_config = config.get("sap", {})

    workbook_path = config["excel"]["path"]
    cost_center = str(sap_config.get("cost_center", "2003568"))
    project_row = int(sap_config.get("project_row", 0))
    url = args.url or sap_config.get("url") or SAP_URL
    profile_path = Path(
        args.profile_path
        or sap_config.get("profile_path")
        or Path(__file__).resolve().parents[2] / "data" / "firefox-sap-profile"
    )

    entries = read_month_entries(workbook_path, args.year, args.month)
    print_plan(entries, cost_center, args.year, args.month)
    if args.browser_js:
        script = browser_script(entries, cost_center, project_row)
        if args.copy:
            import subprocess

            subprocess.run(["clip.exe"], input=script, text=True, check=True)
            print("Browser JavaScript copied to the Windows clipboard.")
            print_browser_console_instructions(copied=True)
        else:
            print("\nPaste this JavaScript into the Firefox console on the SAP form:\n")
            print(script)
            print_browser_console_instructions(copied=False)
        return
    if args.dry_run:
        return

    with sync_playwright() as playwright:
        browser = playwright.firefox.launch_persistent_context(
            user_data_dir=str(profile_path),
            headless=False,
            ignore_https_errors=True,
        )
        page = browser.pages[0] if browser.pages else browser.new_page()
        page.goto(url)
        input("Log in/open the Zeitnachweis form if needed, then press Enter here to fill values...")
        fill_sap_form(page, entries, cost_center, project_row)
        input("Review the SAP form. Press Enter here to close the automation browser...")
        browser.close()


if __name__ == "__main__":
    main()
