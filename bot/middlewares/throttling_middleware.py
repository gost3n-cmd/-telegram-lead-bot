"""Throttling Middleware (Presentation layer).

Защита от флуда — ограничивает частоту сообщений от одного telegram_id.
"""

from __future__ import annotations

import asyncio
from collections import defaultdict
from datetime import datetime, timedelta, timezone
from typing import Any, Awaitable, Callable, Dict

from aiogram import BaseMiddleware
from aiogram.types import Message, TelegramObject

from loguru import logger


class ThrottlingMiddleware(BaseMiddleware):
    """Ограничение частоты сообщений от одного пользователя.

    Parameters:
        max_messages: Максимальное количество сообщений в окне.
        window_seconds: Размер окна в секундах.
    """

    def __init__(self, max_messages: int = 5, window_seconds: int = 10) -> None:
        self._max_messages = max_messages
        self._window = timedelta(seconds=window_seconds)
        self._user_history: Dict[int, list[datetime]] = defaultdict(list)

    async def _cleanup(self, user_id: int, now: datetime) -> None:
        """Удалить устаревшие записи из истории."""
        cutoff = now - self._window
        self._user_history[user_id] = [
            ts for ts in self._user_history[user_id] if ts > cutoff
        ]

    async def __call__(
        self,
        handler: Callable[[TelegramObject, Dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: Dict[str, Any],
    ) -> Any:
        if not isinstance(event, Message) or not event.from_user:
            return await handler(event, data)

        user_id = event.from_user.id
        now = datetime.now(timezone.utc)

        await self._cleanup(user_id, now)
        self._user_history[user_id].append(now)

        if len(self._user_history[user_id]) > self._max_messages:
            logger.warning(
                f"Throttled user_id={user_id}: "
                f"{len(self._user_history[user_id])} messages in {self._window.total_seconds()}s"
            )
            return  # Молча игнорируем — не спамим уведомлениями

        return await handler(event, data)
