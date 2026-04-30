from __future__ import annotations

import asyncio
import json
from copy import deepcopy
from datetime import datetime
from pathlib import Path
from typing import Any


DEFAULT_STATE: dict[str, Any] = {
    "pending_date": None,
    "language": "de",
    "conversation": {
        "mode": "IDLE",
        "target_date": None,
        "start": None,
        "end": None,
    },
    "last_prompted_date": None,
    "last_retry_date": None,
}


class StateStore:
    def __init__(self, path: str) -> None:
        self.path = Path(path)
        self._lock = asyncio.Lock()
        self._state: dict[str, Any] = deepcopy(DEFAULT_STATE)

    async def load(self) -> dict[str, Any]:
        async with self._lock:
            self.path.parent.mkdir(parents=True, exist_ok=True)
            if self.path.exists():
                loaded = json.loads(self.path.read_text(encoding="utf-8"))
                self._state = deepcopy(DEFAULT_STATE)
                self._state.update(loaded)
                self._state["conversation"] = {
                    **deepcopy(DEFAULT_STATE["conversation"]),
                    **loaded.get("conversation", {}),
                }
            else:
                await self._write_locked()
            return deepcopy(self._state)

    async def snapshot(self) -> dict[str, Any]:
        async with self._lock:
            return deepcopy(self._state)

    async def get(self, key: str, default: Any = None) -> Any:
        async with self._lock:
            return deepcopy(self._state.get(key, default))

    async def set(self, key: str, value: Any) -> None:
        async with self._lock:
            self._state[key] = value
            await self._write_locked()

    async def update_conversation(self, **updates: Any) -> dict[str, Any]:
        async with self._lock:
            self._state["conversation"].update(updates)
            await self._write_locked()
            return deepcopy(self._state["conversation"])

    async def reset_conversation(self) -> None:
        async with self._lock:
            self._state["conversation"] = deepcopy(DEFAULT_STATE["conversation"])
            await self._write_locked()

    async def _write_locked(self) -> None:
        self.path.write_text(
            json.dumps(self._state, indent=2, ensure_ascii=True) + "\n",
            encoding="utf-8",
        )
        # Write hermes-status.json for external tools (e.g. Hermes 8am cron)
        status_path = self.path.parent / "hermes-status.json"
        status_path.write_text(
            json.dumps(
                {
                    "pending_date": self._state.get("pending_date"),
                    "last_prompted_date": self._state.get("last_prompted_date"),
                    "conversation_mode": self._state["conversation"]["mode"],
                    "updated_at": datetime.utcnow().isoformat(),
                },
                indent=2,
            ) + "\n",
            encoding="utf-8",
        )
