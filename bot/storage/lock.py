"""Механизм сериализации записи — asyncio.Lock wrapper (Infrastructure layer).

Раздел 14 архитектуры: один писатель за раз.
Лок создаётся как атрибут ExcelStorage при инициализации.
Инстанция ExcelStorage создаётся в main.py и передаётся через DI —
это гарантирует, что в пределах процесса существует ровно один экземпляр лока.
"""

from __future__ import annotations

import asyncio
from contextlib import asynccontextmanager
from typing import AsyncIterator


class AsyncIOLock:
    """Wrapper вокруг asyncio.Lock для сериализации записи в хранилище.

    asyncio.Lock работает кооперативно: второй save_lead отдаёт управление
    event loop и возобновляется после освобождения.
    """

    def __init__(self) -> None:
        self._lock = asyncio.Lock()

    @asynccontextmanager
    async def acquire(self) -> AsyncIterator[None]:
        """Контекстный менеджер для захвата/освобождения лока."""
        async with self._lock:
            yield

    @property
    def locked(self) -> bool:
        return self._lock.locked()
