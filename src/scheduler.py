from __future__ import annotations

from apscheduler.schedulers.asyncio import AsyncIOScheduler
from apscheduler.triggers.cron import CronTrigger


class BotScheduler:
    def __init__(self, timezone: str, schedule_config: dict, conversation) -> None:
        self.conversation = conversation
        self.scheduler = AsyncIOScheduler(timezone=timezone)
        daily_hour, daily_minute = self._parse_clock(schedule_config["daily_prompt"])
        retry_hour, retry_minute = self._parse_clock(schedule_config["morning_retry"])

        self.scheduler.add_job(
            self.conversation.handle_daily_prompt,
            CronTrigger(
                day_of_week="mon-fri",
                hour=daily_hour,
                minute=daily_minute,
                timezone=timezone,
            ),
            id="daily_prompt",
            replace_existing=True,
            misfire_grace_time=900,
        )
        self.scheduler.add_job(
            self.conversation.handle_retry_prompt,
            CronTrigger(
                day_of_week="mon-fri",
                hour=retry_hour,
                minute=retry_minute,
                timezone=timezone,
            ),
            id="retry_prompt",
            replace_existing=True,
            misfire_grace_time=900,
        )

    async def start(self) -> None:
        self.scheduler.start()

    async def stop(self) -> None:
        if self.scheduler.running:
            self.scheduler.shutdown(wait=False)

    @staticmethod
    def _parse_clock(value: str) -> tuple[int, int]:
        hour_str, minute_str = value.split(":", 1)
        return int(hour_str), int(minute_str)
