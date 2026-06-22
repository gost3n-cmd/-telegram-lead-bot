"""StorageInterface — абстрактный контракт хранилища (Domain layer).

Определён в разделе 10 архитектуры.
Application (lead_service) зависит от этого интерфейса, а не от конкретной реализации.
Все методы оперируют только примитивами и моделью Lead — ничего специфичного для Excel.
"""

from __future__ import annotations

from abc import ABC, abstractmethod
from typing import List, Optional

from bot.models.lead import Lead


class SaveResult:
    """Результат операции сохранения лида.

    Аналог `NamedTuple` для совместимости с pydantic-free подходом,
    но с возможностью расширения в будущем.
    """

    __slots__ = ("success", "error", "is_duplicate")

    def __init__(self, success: bool, error: Optional[str] = None, is_duplicate: bool = False) -> None:
        self.success = success
        self.error = error
        self.is_duplicate = is_duplicate

    def __bool__(self) -> bool:
        return self.success

    def __repr__(self) -> str:
        if self.error:
            return f"SaveResult(success={self.success}, error={self.error!r})"
        return f"SaveResult(success={self.success})"


class StorageInterface(ABC):
    """Абстрактный интерфейс хранилища лидов.

    Контракт, от которого зависит lead_service.
    Конкретные реализации (Excel, SQLite, Postgres) реализуют этот интерфейс.
    """

    @abstractmethod
    async def save_lead(self, lead: Lead) -> SaveResult:
        """Атомарно добавить одну запись в хранилище."""
        ...

    @abstractmethod
    async def email_exists(self, email: str) -> bool:
        """Проверить уникальность email перед сохранением."""
        ...

    @abstractmethod
    async def count_leads(self) -> int:
        """Текущее количество сохранённых записей — для статистики/админ-команд."""
        ...

    @abstractmethod
    async def list_leads(self, limit: int = 50, offset: int = 0) -> List[Lead]:
        """Постраничная выборка — задел под будущую админ-панель или экспорт."""
        ...

    @abstractmethod
    async def health_check(self) -> bool:
        """Проверка доступности хранилища (файл не залочен / БД отвечает)."""
        ...
