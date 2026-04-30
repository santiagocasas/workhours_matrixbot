from __future__ import annotations

import logging
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

from src.excel.time_utils import (
    format_duration_hours,
    parse_break_ranges,
    parse_date_input,
    parse_time_input,
)


SPECIAL_CODES = {"k": "K", "u": "U", "g": "G"}
SUPPORTED_LANGUAGES = {"de", "en"}

STRINGS = {
    "de": {
        "help": (
            "Befehle:\n"
            "!help\n"
            "!today\n"
            "!missed [TT.MM]\n"
            "!status [TT.MM oder TT.MM.JJJJ]\n"
            "!correct TT.MM <k|u|g>\n"
            "!correct TT.MM <Start> <Ende> [Pause1,Pause2]\n"
            "!language [de|en]\n"
            "!testreminder"
        ),
        "weekend_today": "Heute ist Wochenende. Falls noetig, nutze !correct fuer ein konkretes Datum.",
        "today_auto_skip": "Heute ({date}) ist {reason}.",
        "today_has_entry": "Heute bereits eingetragen: {status} Zum Aendern nutze !correct.",
        "unknown_command": "Unbekannter Befehl. Mit !help bekommst du die verfuegbaren Befehle.",
        "idle_no_prompt": "Keine offene Abfrage. Mit !today kannst du die heutige Eingabe starten.",
        "retry_prompt": "Erinnerung fuer {date}: Hast du gearbeitet? (yes / ja / k / u / g / skip)",
        "daily_prompt": "Guten Tag! Eintrag fuer {date}: Hast du gearbeitet? (yes / ja / k / u / g / skip)",
        "correction_usage": "Format: !correct TT.MM <k|u|g> oder !correct TT.MM 09:00 17:30 [12:00-12:30]",
        "missing_start_end": "Es fehlen Start- und Endzeit.",
        "corrected": "Korrigiert: {summary}",
        "skipped": "Ok, uebersprungen.",
        "worked_invalid": "Bitte antworte mit yes, ja, k, u, g oder skip.",
        "missed_none": "Keine fehlenden Eintraege in den letzten 14 Tagen.",
        "missed_list": "Fehlende Eintraege:\n{dates}",
        "ask_start": "Um wie viel Uhr hast du angefangen? (z.B. 09:00)",
        "ask_end": "Um wie viel Uhr hast du aufgehoert? (z.B. 17:30)",
        "ask_breaks": "Gab es Unterbrechungen? (z.B. 12:00-12:30, 15:00-15:15 oder 'nein')",
        "status_auto_skip": "{date} ist {reason}.",
        "status_empty": "{date}: noch nichts eingetragen.",
        "status_special": "{date}: Schluessel {code}.",
        "status_workday": "{date}: {start} - {end}, Pausen: {breaks}, erwartete Stunden: {hours}.",
        "write_special": "{date}: Schluessel {code} eingetragen.",
        "write_workday": "{date}: {start}-{end}, Pausen: {breaks}, erwartete Stunden: {hours}.",
        "copy_note": "(Datei auch nach Windows kopiert)",
        "language_current": "Aktuelle Sprache: {language}.",
        "language_set": "Sprache gesetzt auf {language}.",
        "language_invalid": "Bitte nutze !language de oder !language en.",
        "language_name_de": "Deutsch",
        "language_name_en": "Englisch",
    },
    "en": {
        "help": (
            "Commands:\n"
            "!help\n"
            "!today\n"
            "!missed [DD.MM]\n"
            "!status [DD.MM or DD.MM.YYYY]\n"
            "!correct DD.MM <k|u|g>\n"
            "!correct DD.MM <start> <end> [break1,break2]\n"
            "!language [de|en]\n"
            "!testreminder"
        ),
        "weekend_today": "Today is a weekend. If needed, use !correct for a specific date.",
        "today_auto_skip": "Today ({date}) is {reason}.",
        "today_has_entry": "Today is already entered: {status} Use !correct to change it.",
        "unknown_command": "Unknown command. Use !help to see the available commands.",
        "idle_no_prompt": "No active prompt. Use !today to start today's entry.",
        "retry_prompt": "Reminder for {date}: Did you work? (yes / ja / k / u / g / skip)",
        "daily_prompt": "Hello! Entry for {date}: Did you work? (yes / ja / k / u / g / skip)",
        "correction_usage": "Format: !correct DD.MM <k|u|g> or !correct DD.MM 09:00 17:30 [12:00-12:30]",
        "missing_start_end": "Start and end times are missing.",
        "corrected": "Updated: {summary}",
        "skipped": "Ok, skipped.",
        "worked_invalid": "Please reply with yes, ja, k, u, g or skip.",
        "missed_none": "No missing entries in the last 14 days.",
        "missed_list": "Missing entries:\n{dates}",
        "ask_start": "What time did you start? (for example 09:00)",
        "ask_end": "What time did you finish? (for example 17:30)",
        "ask_breaks": "Were there breaks? (for example 12:00-12:30, 15:00-15:15 or 'no')",
        "status_auto_skip": "{date} is {reason}.",
        "status_empty": "{date}: nothing entered yet.",
        "status_special": "{date}: code {code}.",
        "status_workday": "{date}: {start} - {end}, breaks: {breaks}, expected hours: {hours}.",
        "write_special": "{date}: code {code} entered.",
        "write_workday": "{date}: {start}-{end}, breaks: {breaks}, expected hours: {hours}.",
        "copy_note": "(File was also copied to Windows)",
        "language_current": "Current language: {language}.",
        "language_set": "Language set to {language}.",
        "language_invalid": "Please use !language de or !language en.",
        "language_name_de": "German",
        "language_name_en": "English",
    },
}

AUTO_REASON_KEYS = {
    "Wochenende": {"de": "Wochenende", "en": "weekend"},
    "Feiertag": {"de": "Feiertag", "en": "holiday"},
    "Brueckentag": {"de": "Brueckentag", "en": "bridge day"},
}

LOGGER = logging.getLogger(__name__)


class ConversationManager:
    def __init__(self, config: dict, state, excel, send_text) -> None:
        self.config = config
        self.state = state
        self.excel = excel
        self.send_text = send_text
        self.tz = ZoneInfo(config["timezone"])
        self.lang = "de"

    async def handle_startup_catchup(self) -> None:
        """Send the daily prompt at startup if it was missed today after prompt time."""
        await self.handle_periodic_catchup(reason="startup")

    async def handle_periodic_catchup(self, reason: str = "scheduled") -> None:
        """Recover prompts missed while the machine or WSL session was asleep."""
        today = self._today()
        now = datetime.now(self.tz)
        state = await self.state.snapshot()
        self.lang = self._normalize_language(state.get("language"))

        if state.get("pending_date"):
            LOGGER.info(
                "Catch-up skipped: pending prompt already exists for %s",
                state["pending_date"],
            )
            return

        if today.weekday() >= 5:
            LOGGER.info("Catch-up skipped: today is weekend")
            return

        daily_time = self.config["schedule"]["daily_prompt"]
        prompt_hour, prompt_minute = [int(x) for x in daily_time.split(":", 1)]
        retry_time = self.config["schedule"]["morning_retry"]
        retry_hour, retry_minute = [int(x) for x in retry_time.split(":", 1)]
        before_daily_prompt = now.hour < prompt_hour or (
            now.hour == prompt_hour and now.minute < prompt_minute
        )
        before_retry_prompt = now.hour < retry_hour or (
            now.hour == retry_hour and now.minute < retry_minute
        )
        if now.hour < prompt_hour or (now.hour == prompt_hour and now.minute < prompt_minute):
            LOGGER.info("Catch-up skipped today: before daily prompt time")

        last_prompted = state.get("last_prompted_date")
        today_iso = today.isoformat()
        if not before_daily_prompt and last_prompted != today_iso:
            status = self.excel.get_status(today)
            if status["is_auto_skip"]:
                LOGGER.info("Catch-up skipped: today is %s", status["auto_reason"])
            elif status["has_entry"]:
                LOGGER.info("Catch-up skipped: today already has an entry")
            else:
                LOGGER.info("Catch-up sending prompt for today (%s): %s", today_iso, reason)
                await self._start_prompt(today, is_retry=False)
                return
        elif last_prompted == today_iso:
            LOGGER.info("Catch-up skipped: today already prompted")

        yesterday = today - timedelta(days=1)
        if yesterday.weekday() >= 5:
            LOGGER.info("Catch-up skipped: yesterday was weekend")
            return
        if before_retry_prompt:
            LOGGER.info("Catch-up skipped yesterday: before retry prompt time")
            return

        yesterday_iso = yesterday.isoformat()
        y_status = self.excel.get_status(yesterday)
        if y_status["is_auto_skip"]:
            LOGGER.info("Catch-up skipped: yesterday was %s", y_status["auto_reason"])
            return
        if y_status["has_entry"]:
            LOGGER.info("Catch-up skipped: yesterday already has an entry")
            return

        LOGGER.info("Catch-up sending prompt for yesterday (%s): %s", yesterday_iso, reason)
        await self._start_prompt(
            yesterday,
            is_retry=False,
            prefix="[Yesterday catch-up] ",
        )

    async def handle_daily_prompt(self) -> None:
        today = self._today()
        if today.weekday() >= 5:
            return
        if await self._date_should_be_skipped(today):
            return
        await self._start_prompt(today, is_retry=False)

    async def handle_retry_prompt(self) -> None:
        state = await self.state.snapshot()
        self.lang = self._normalize_language(state.get("language"))
        pending_raw = state.get("pending_date")
        if not pending_raw:
            return

        pending_date = date.fromisoformat(pending_raw)
        today = self._today()
        if pending_date != today - timedelta(days=1):
            return

        last_retry = state.get("last_retry_date")
        if last_retry == today.isoformat():
            return

        if await self._date_should_be_skipped(pending_date):
            await self.state.set("pending_date", None)
            return

        await self.send_text(
            self._text("retry_prompt", date=pending_date.strftime("%d.%m.%Y"))
        )
        await self.state.update_conversation(
            mode="ASKED_WORKED",
            target_date=pending_date.isoformat(),
            start=None,
            end=None,
        )
        await self.state.set("last_retry_date", today.isoformat())

    async def handle_message(self, text: str) -> None:
        snapshot = await self.state.snapshot()
        self.lang = self._normalize_language(snapshot.get("language"))
        try:
            message = text.strip()
            if not message:
                return
            if message.startswith("!"):
                await self._handle_command(message)
                return

            conversation = snapshot["conversation"]
            mode = conversation["mode"]
            if mode == "IDLE":
                await self.send_text(self._text("idle_no_prompt"))
                return

            target_date = date.fromisoformat(conversation["target_date"])
            lowered = message.lower()

            if mode == "ASKED_WORKED":
                await self._handle_worked_answer(target_date, lowered)
                return
            if mode == "ASKED_START":
                await self._handle_start_answer(message)
                return
            if mode == "ASKED_END":
                await self._handle_end_answer(message)
                return
            if mode == "ASKED_BREAKS":
                await self._handle_break_answer(target_date, message)
                return
        except ValueError as exc:
            await self.send_text(str(exc))

    async def _handle_command(self, command: str) -> None:
        try:
            parts = command.split(maxsplit=1)
            name = parts[0].lower()
            rest = parts[1] if len(parts) > 1 else ""

            if name == "!help":
                await self.send_text(self._text("help"))
                return

            if name == "!language":
                await self._handle_language_command(rest)
                return

            if name == "!today":
                today = self._today()
                if today.weekday() >= 5:
                    await self.send_text(self._text("weekend_today"))
                    return
                status = self.excel.get_status(today)
                if status["is_auto_skip"]:
                    await self.send_text(
                        self._text(
                            "today_auto_skip",
                            date=today.strftime("%d.%m.%Y"),
                            reason=self._translate_auto_reason(status["auto_reason"]),
                        )
                    )
                    return
                if status["has_entry"]:
                    await self.send_text(
                        self._text(
                            "today_has_entry",
                            status=self._format_status_message(today, status),
                        )
                    )
                    return
                await self._start_prompt(today, is_retry=False)
                return

            if name == "!testreminder":
                await self._start_prompt(
                    self._today(), is_retry=False, prefix="[TEST] "
                )
                return

            if name == "!missed":
                await self._handle_missed_command(rest)
                return

            if name == "!status":
                target_date = parse_date_input(
                    rest or self._today().strftime("%d.%m.%Y"), self._today().year
                )
                status = self.excel.get_status(target_date)
                await self.send_text(self._format_status_message(target_date, status))
                return

            if name == "!correct":
                await self._handle_correction(rest)
                return

            await self.send_text(self._text("unknown_command"))
        except ValueError as exc:
            await self.send_text(str(exc))

    async def _handle_missed_command(self, payload: str) -> None:
        payload = payload.strip()
        if payload:
            # Specific date provided: start prompt for that date
            target_date = parse_date_input(payload, self._today().year)
            await self._start_prompt(target_date, is_retry=False)
            return
        # No date: list all missing weekdays in past 14 days
        today = self._today()
        missing = []
        for days_back in range(1, 15):
            check_date = today - timedelta(days=days_back)
            if check_date.weekday() >= 5:
                continue
            status = self.excel.get_status(check_date)
            if not status["is_auto_skip"] and not status["has_entry"]:
                missing.append(check_date.strftime("%d.%m.%Y (%A)"))
        if not missing:
            await self.send_text(self._text("missed_none"))
        else:
            await self.send_text(self._text("missed_list", dates="\n".join(missing)))

    async def _handle_language_command(self, payload: str) -> None:
        selection = payload.strip().lower()
        if not selection:
            await self.send_text(
                self._text("language_current", language=self._language_name(self.lang))
            )
            return
        if selection not in SUPPORTED_LANGUAGES:
            await self.send_text(self._text("language_invalid"))
            return

        self.lang = selection
        await self.state.set("language", selection)
        await self.send_text(
            self._text("language_set", language=self._language_name(selection))
        )

    async def _handle_correction(self, payload: str) -> None:
        tokens = payload.split()
        if len(tokens) < 2:
            await self.send_text(self._text("correction_usage"))
            return

        target_date = parse_date_input(tokens[0], self._today().year)
        action = tokens[1].lower()

        if action in SPECIAL_CODES:
            result = self.excel.write_special_day(target_date, SPECIAL_CODES[action])
            await self.send_text(
                self._with_copy_note(
                    self._text("corrected", summary=self._format_write_result(result)),
                    result,
                )
            )
            return

        if len(tokens) < 3:
            await self.send_text(self._text("missing_start_end"))
            return

        start = parse_time_input(tokens[1])
        end = parse_time_input(tokens[2])
        breaks_text = " ".join(tokens[3:]).strip()
        breaks = parse_break_ranges(breaks_text) if breaks_text else []
        result = self.excel.write_workday(target_date, start, end, breaks)
        await self.send_text(
            self._with_copy_note(
                self._text("corrected", summary=self._format_write_result(result)),
                result,
            )
        )

    async def _handle_worked_answer(self, target_date: date, answer: str) -> None:
        if answer == "skip":
            await self.state.reset_conversation()
            await self.send_text(self._text("skipped"))
            return
        if answer in SPECIAL_CODES:
            result = self.excel.write_special_day(target_date, SPECIAL_CODES[answer])
            await self.state.reset_conversation()
            await self.state.set("pending_date", None)
            await self.send_text(
                self._with_copy_note(self._format_write_result(result), result)
            )
            return
        if answer not in {"yes", "ja"}:
            await self.send_text(self._text("worked_invalid"))
            return

        await self.state.update_conversation(mode="ASKED_START")
        await self.send_text(self._text("ask_start"))

    async def _handle_start_answer(self, message: str) -> None:
        start = parse_time_input(message)
        await self.state.update_conversation(
            mode="ASKED_END", start=self._serialize_timedelta(start)
        )
        await self.send_text(self._text("ask_end"))

    async def _handle_end_answer(self, message: str) -> None:
        end = parse_time_input(message)
        await self.state.update_conversation(
            mode="ASKED_BREAKS", end=self._serialize_timedelta(end)
        )
        await self.send_text(self._text("ask_breaks"))

    async def _handle_break_answer(self, target_date: date, message: str) -> None:
        snapshot = await self.state.snapshot()
        conversation = snapshot["conversation"]
        start = self._deserialize_timedelta(conversation["start"])
        end = self._deserialize_timedelta(conversation["end"])
        breaks = parse_break_ranges(message)
        result = self.excel.write_workday(target_date, start, end, breaks)

        await self.state.reset_conversation()
        await self.state.set("pending_date", None)
        await self.send_text(
            self._with_copy_note(self._format_write_result(result), result)
        )

    async def _date_should_be_skipped(self, target_date: date) -> bool:
        status = self.excel.get_status(target_date)
        return status["is_auto_skip"] or status["has_entry"]

    async def _start_prompt(
        self, target_date: date, is_retry: bool, prefix: str = ""
    ) -> None:
        await self.state.set("pending_date", target_date.isoformat())
        await self.state.update_conversation(
            mode="ASKED_WORKED",
            target_date=target_date.isoformat(),
            start=None,
            end=None,
        )
        await self.state.set("last_prompted_date", target_date.isoformat())
        if not is_retry:
            await self.send_text(
                prefix
                + self._text("daily_prompt", date=target_date.strftime("%d.%m.%Y"))
            )

    def _format_status_message(self, target_date: date, status: dict) -> str:
        if status["is_auto_skip"]:
            return self._text(
                "status_auto_skip",
                date=target_date.strftime("%d.%m.%Y"),
                reason=self._translate_auto_reason(status["auto_reason"]),
            )
        if not status["has_entry"]:
            return self._text("status_empty", date=target_date.strftime("%d.%m.%Y"))
        if status["special_code"]:
            return self._text(
                "status_special",
                date=target_date.strftime("%d.%m.%Y"),
                code=status["special_code"],
            )
        breaks = status["breaks_text"] or ("keine" if self.lang == "de" else "none")
        preview = status.get("preview_hours")
        return self._text(
            "status_workday",
            date=target_date.strftime("%d.%m.%Y"),
            start=status["start_text"],
            end=status["end_text"],
            breaks=breaks,
            hours=format_duration_hours(preview),
        )

    def _text(self, key: str, **kwargs) -> str:
        template = STRINGS[self.lang][key]
        return template.format(**kwargs)

    def _format_write_result(self, result) -> str:
        date_text = result.target_date.strftime("%d.%m.%Y")
        if result.kind == "special":
            return self._text("write_special", date=date_text, code=result.code)

        breaks = result.breaks_text or ("keine" if self.lang == "de" else "none")
        return self._text(
            "write_workday",
            date=date_text,
            start=result.start_text,
            end=result.end_text,
            breaks=breaks,
            hours=format_duration_hours(result.preview_hours),
        )

    def _translate_auto_reason(self, reason: str | None) -> str:
        if reason is None:
            return "-"
        mapping = AUTO_REASON_KEYS.get(reason)
        if mapping:
            return mapping[self.lang]
        return reason

    def _language_name(self, code: str) -> str:
        return STRINGS[self.lang][f"language_name_{code}"]

    @staticmethod
    def _normalize_language(value: str | None) -> str:
        if value in SUPPORTED_LANGUAGES:
            return value
        return "de"

    def _with_copy_note(self, message: str, result) -> str:
        if getattr(result, "windows_copied", False):
            return message + "\n" + self._text("copy_note")
        return message

    def _today(self) -> date:
        return datetime.now(self.tz).date()

    @staticmethod
    def _serialize_timedelta(value: timedelta) -> int:
        return int(value.total_seconds())

    @staticmethod
    def _deserialize_timedelta(value: int | None) -> timedelta:
        if value is None:
            raise ValueError("Missing time value in conversation state")
        return timedelta(seconds=value)
