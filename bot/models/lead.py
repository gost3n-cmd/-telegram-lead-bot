"""Lead — модель данных одной заявки (Domain layer).

Поля определены в разделе 10 архитектуры.
Лист без исходящих зависимостей внутри проекта.
"""

from __future__ import annotations

from datetime import datetime, timezone
from typing import Optional

from pydantic import BaseModel, Field


class Lead(BaseModel):
    """Структура данных одной заявки на подписку."""

    telegram_id: int = Field(..., description="Уникальный идентификатор пользователя в Telegram")
    telegram_username: Optional[str] = Field(default=None, description="Username, может отсутствовать")
    email: str = Field(..., description="Нормализованный email (lowercase, без пробелов)")
    submitted_at: datetime = Field(
        default_factory=lambda: datetime.now(timezone.utc),
        description="Момент сохранения (UTC)",
    )
    source: str = Field(default="bot_v1", description="Источник заявки, задел на несколько каналов")

    model_config = {"frozen": True}
