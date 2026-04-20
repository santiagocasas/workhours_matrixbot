from __future__ import annotations

import asyncio
import logging
import sys
from pathlib import Path
from typing import Any

import yaml

from src.bot.conversation import ConversationManager
from src.bot.matrix_client import MatrixBotClient
from src.excel.handler import ExcelHandler
from src.scheduler import BotScheduler
from src.state import StateStore


logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    stream=sys.stdout,
)
LOGGER = logging.getLogger("work-hours-bot")


def load_config() -> dict[str, Any]:
    config_path = Path(__file__).resolve().parent.parent / "config.yaml"
    with config_path.open("r", encoding="utf-8") as handle:
        return yaml.safe_load(handle)


async def main() -> None:
    config = load_config()

    LOGGER.info("Starting work-hours-bot")
    LOGGER.info("Bot user: %s", config["matrix"]["bot_user_id"])
    LOGGER.info("Room: %s", config["matrix"]["room_id"])
    LOGGER.info("Allowed user: %s", config["matrix"]["allowed_user_id"])
    LOGGER.info("Excel path: %s", config["excel"]["path"])

    state = StateStore(config["state_file"])
    await state.load()
    LOGGER.info("State loaded from %s", config["state_file"])

    excel = ExcelHandler(
        config["excel"]["path"],
        windows_path=config["excel"].get("windows_path"),
    )
    matrix = MatrixBotClient(
        homeserver=config["matrix"]["homeserver"],
        bot_user_id=config["matrix"]["bot_user_id"],
        bot_password=config["matrix"]["bot_password"],
        room_id=config["matrix"]["room_id"],
        allowed_user_id=config["matrix"]["allowed_user_id"],
    )

    conversation = ConversationManager(
        config=config,
        state=state,
        excel=excel,
        send_text=matrix.send_text,
    )

    scheduler = BotScheduler(
        timezone=config["timezone"],
        schedule_config=config["schedule"],
        conversation=conversation,
    )

    LOGGER.info("Logging into Matrix...")
    await matrix.login()
    LOGGER.info("Login successful. Starting scheduler...")
    await scheduler.start()
    await conversation.handle_startup_catchup()
    LOGGER.info("Startup catchup check done.")
    LOGGER.info(
        "Scheduler started. Daily prompt at %s, retry at %s. Listening for messages...",
        config["schedule"]["daily_prompt"],
        config["schedule"]["morning_retry"],
    )

    try:
        await matrix.run(conversation.handle_message)
    finally:
        await scheduler.stop()
        await matrix.close()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        LOGGER.info("Stopped by user.")
    except Exception:
        LOGGER.exception("Fatal error")
        sys.exit(1)
