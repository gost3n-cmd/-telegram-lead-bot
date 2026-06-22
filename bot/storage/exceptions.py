"""Типизированные исключения хранилища (Infrastructure layer).

Раздел 15 архитектуры: StorageError, StorageUnavailableError.
"""

from __future__ import annotations


class StorageError(Exception):
    """Базовое исключение хранилища — непредвиденная ошибка при записи."""

    def __init__(self, message: str = "Storage error") -> None:
        self.message = message
        super().__init__(self.message)


class StorageUnavailableError(StorageError):
    """Хранилище недоступно: файл занят, нет прав, нет места, повреждён."""

    def __init__(self, message: str = "Storage unavailable") -> None:
        super().__init__(message)


class DuplicateLeadError(StorageError):
    """Дубликат email — попытка сохранить уже существующий лид.

    Примечание: в текущей архитектуре дубликаты проверяются через
    email_exists() в lead_service, а не через исключение.
    Этот класс оставлен как Infrastructure-level safeguard.
    """

    def __init__(self, email: str) -> None:
        self.email = email
        super().__init__(f"Duplicate lead: {email}")
