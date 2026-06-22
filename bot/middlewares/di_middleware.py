"""DI Middleware — внедрение зависимостей в handler-контекст (Presentation layer).

Передаёт зарегистрированные сервисы (lead_service) в kwargs хендлеров.
Это позволяет хендлерам получать зависимости через параметры функции,
без прямого импорта конкретных реализаций.
"""

from __future__ import annotations

from typing import Any, Awaitable, Callable, Dict

from aiogram import BaseMiddleware
from aiogram.types import TelegramObject


class DIMiddleware(BaseMiddleware):
    """Внедряет зависимости из Dispatcher data в handler kwargs."""

    async def __call__(
        self,
        handler: Callable[[TelegramObject, Dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: Dict[str, Any],
    ) -> Any:
        # Передаём все зарегистрированные сервисы в handler kwargs
        return await handler(event, data)
