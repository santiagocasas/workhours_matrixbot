from __future__ import annotations

import logging
from typing import Awaitable, Callable

from nio import AsyncClient, LoginResponse, MatrixRoom, RoomMessageText


LOGGER = logging.getLogger(__name__)


class MatrixBotClient:
    def __init__(
        self,
        homeserver: str,
        bot_user_id: str,
        bot_password: str,
        room_id: str,
        allowed_user_id: str,
    ) -> None:
        self.room_id = room_id
        self.allowed_user_id = allowed_user_id
        self.client = AsyncClient(homeserver, bot_user_id)
        self._password = bot_password
        self._message_handler: Callable[[str], Awaitable[None]] | None = None

    async def login(self) -> None:
        response = await self.client.login(self._password)
        if not isinstance(response, LoginResponse):
            raise RuntimeError(f"Matrix login failed: {response}")
        # Prime the sync token so the bot does not react to older room history.
        await self.client.sync(timeout=3_000, full_state=True)

    async def run(self, message_handler: Callable[[str], Awaitable[None]]) -> None:
        self._message_handler = message_handler
        self.client.add_event_callback(self._on_message, RoomMessageText)
        await self.client.sync_forever(timeout=30_000, full_state=True)

    async def send_text(self, text: str) -> None:
        await self.client.room_send(
            room_id=self.room_id,
            message_type="m.room.message",
            content={
                "msgtype": "m.text",
                "body": text,
            },
        )

    async def close(self) -> None:
        await self.client.close()

    async def _on_message(self, room: MatrixRoom, event: RoomMessageText) -> None:
        if room.room_id != self.room_id:
            return
        if event.sender != self.allowed_user_id:
            return
        if event.sender == self.client.user_id:
            return
        if self._message_handler is None:
            LOGGER.warning("Matrix message received without handler")
            return
        await self._message_handler(event.body.strip())
