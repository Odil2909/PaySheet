from __future__ import annotations

import asyncio
import json
from typing import Any, Mapping

from aiogram.fsm.state import State
from aiogram.fsm.storage.base import BaseStorage, StorageKey

import database


class SQLiteFSMStorage(BaseStorage):
    @staticmethod
    def _key(key: StorageKey) -> str:
        return json.dumps(
            [
                key.bot_id,
                key.chat_id,
                key.user_id,
                key.thread_id,
                key.business_connection_id,
                key.destiny,
            ],
            separators=(",", ":"),
        )

    async def set_state(
        self, key: StorageKey, state: str | State | None = None
    ) -> None:
        value = state.state if isinstance(state, State) else state
        await asyncio.to_thread(database.save_fsm_state, self._key(key), value)

    async def get_state(self, key: StorageKey) -> str | None:
        state, _ = await asyncio.to_thread(database.load_fsm, self._key(key))
        return state

    async def set_data(self, key: StorageKey, data: Mapping[str, Any]) -> None:
        await asyncio.to_thread(
            database.save_fsm_data, self._key(key), dict(data)
        )

    async def get_data(self, key: StorageKey) -> dict[str, Any]:
        _, data = await asyncio.to_thread(database.load_fsm, self._key(key))
        return data

    async def close(self) -> None:
        return None
