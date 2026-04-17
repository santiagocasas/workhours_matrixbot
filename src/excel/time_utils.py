from __future__ import annotations

from datetime import date, datetime, time, timedelta


SPECIAL_SOLL_CODES = {"b", "k", "h", "t", "u", "z"}


def parse_time_input(value: str) -> timedelta:
    cleaned = value.strip()
    try:
        parsed = datetime.strptime(cleaned, "%H:%M")
    except ValueError as exc:
        raise ValueError(f"Ungueltige Uhrzeit: {value}. Erwartet HH:MM.") from exc
    return timedelta(hours=parsed.hour, minutes=parsed.minute)


def parse_break_ranges(value: str) -> list[tuple[timedelta, timedelta]]:
    cleaned = value.strip()
    if not cleaned or cleaned.lower() in {"nein", "none", "keine", "no"}:
        return []

    breaks: list[tuple[timedelta, timedelta]] = []
    for item in cleaned.split(","):
        part = item.strip()
        if not part:
            continue
        try:
            start_raw, end_raw = part.split("-", 1)
        except ValueError as exc:
            raise ValueError(
                f"Ungueltige Pause: {part}. Erwartet HH:MM-HH:MM."
            ) from exc
        start = parse_time_input(start_raw)
        end = parse_time_input(end_raw)
        breaks.append((start, end))

    if len(breaks) > 3:
        raise ValueError("Es sind maximal 3 Unterbrechungen moeglich.")
    return breaks


def parse_date_input(value: str, default_year: int) -> date:
    cleaned = value.strip()
    for fmt in ("%d.%m.%Y", "%d.%m"):
        try:
            parsed = datetime.strptime(cleaned, fmt)
            year = parsed.year if fmt == "%d.%m.%Y" else default_year
            return date(year, parsed.month, parsed.day)
        except ValueError:
            continue
    raise ValueError(f"Ungueltiges Datum: {value}. Erwartet TT.MM oder TT.MM.JJJJ.")


def duration_to_text(value: timedelta | None) -> str | None:
    if value is None:
        return None
    seconds = int(value.total_seconds())
    hours = (seconds // 3600) % 24
    minutes = (seconds % 3600) // 60
    return f"{hours:02d}:{minutes:02d}"


def format_duration_hours(value: float | None) -> str:
    if value is None:
        return "-"
    return f"{value:.1f}"


def normalize_excel_time(value) -> timedelta | None:
    if value is None:
        return None
    if isinstance(value, timedelta):
        return value
    if isinstance(value, datetime):
        return timedelta(hours=value.hour, minutes=value.minute, seconds=value.second)
    if isinstance(value, time):
        return timedelta(hours=value.hour, minutes=value.minute, seconds=value.second)
    if isinstance(value, (int, float)):
        return timedelta(days=float(value))
    return None


def compute_net_work_duration(
    start: timedelta, end: timedelta, breaks: list[tuple[timedelta, timedelta]]
) -> timedelta:
    gross = end - start
    if gross.total_seconds() < 0:
        gross += timedelta(days=1)

    break_total = timedelta()
    for break_start, break_end in breaks:
        duration = break_end - break_start
        if duration.total_seconds() < 0:
            duration += timedelta(days=1)
        break_total += duration
    return gross - break_total


def compute_preview_hours(
    start: timedelta | None,
    end: timedelta | None,
    breaks: list[tuple[timedelta, timedelta]],
    special_code: str | None,
    soll_duration: timedelta,
) -> float | None:
    code = (special_code or "").strip().lower()
    if code in SPECIAL_SOLL_CODES:
        return round(soll_duration.total_seconds() / 3600, 1)
    if code == "g":
        return 0.0
    if start is None or end is None:
        return None

    gross = end - start
    if gross.total_seconds() < 0:
        gross += timedelta(days=1)

    break_total = timedelta()
    for break_start, break_end in breaks:
        pause = break_end - break_start
        if pause.total_seconds() < 0:
            pause += timedelta(days=1)
        break_total += pause

    net = gross - break_total
    net_fraction = net.total_seconds() / 86400

    if net_fraction <= 6 / 24:
        deduction = 0.0
    elif net_fraction <= 6.5 / 24:
        deduction = net_fraction - 6 / 24
    elif net_fraction <= 9.5 / 24:
        deduction = 0.5 / 24
    elif net_fraction <= 9.75 / 24:
        deduction = net_fraction - 9 / 24
    else:
        deduction = 0.75 / 24

    effective = min(net_fraction - deduction, 10 / 24)
    return round(effective * 24, 1)
