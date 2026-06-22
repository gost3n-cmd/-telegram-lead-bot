"""Logging Middleware (Presentation layer).

Логирует входящие апдейты для отладки и аудита.
"""

from __future__ import annotations

from typing import Any, Awaitable, Callable, Dict

from aiogram import BaseMiddleware
from aiogram.types import TelegramObject

from loguru import logger


class LoggingMiddleware(BaseMiddleware):
    """Логирование входящих обновлений Telegram."""

    async def __call__(
        self,
        handler: Callable[[TelegramObject, Dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: Dict[str, Any],
    ) -> Any:
        # Логируем тип события
        event_type = type(event).__name__
        logger.debug(f"Received event: {event_type}")

        # Для Message — логируем текст (если есть)
        if hasattr(event, "message") and hasattr(event.message, "from_user"):
            user = event.message.from_user
            text = getattr(event.message, "text", None)
            logger.debug(
                f"Message from user_id={user.id} username={user.username}: "
                f"text={text!r:.100}"
            )

        return await handler(event, data)
