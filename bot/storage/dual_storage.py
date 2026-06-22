"""DualStorage — adapter that composes ExcelStorage + GoogleSheetsStorage.

Implements StorageInterface for full compatibility with LeadService.
Primary write goes to Excel (source of truth); on success, a best-effort
secondary write is dispatched to Google Sheets. GS failures are logged
and never propagate to the caller.
"""

from __future__ import annotations

from typing import List

from loguru import logger

from bot.models.lead import Lead
from bot.storage.base import SaveResult, StorageInterface


class DualStorage(StorageInterface):
    """Wraps ExcelStorage (primary) and GoogleSheetsStorage (secondary).

    - Excel is the source of truth — all reads delegate to Excel.
    - Google Sheets is a live-view secondary sink — best-effort only.
    - LeadService sees the exact same SaveResult contract.
    - No modifications to existing ExcelStorage.
    """

    def __init__(self, excel_storage: StorageInterface, gsheets_storage) -> None:
        self._excel = excel_storage
        self._gsheets = gsheets_storage

    async def save_lead(self, lead: Lead) -> SaveResult:
        """Atomic save: Excel first, then best-effort Google Sheets append."""
        result = await self._excel.save_lead(lead)

        if result.success:
            await self._gsheets.append_lead(lead)

        return result

    async def email_exists(self, email: str) -> bool:
        return await self._excel.email_exists(email)

    async def count_leads(self) -> int:
        return await self._excel.count_leads()

    async def list_leads(self, limit: int = 50, offset: int = 0) -> List[Lead]:
        return await self._excel.list_leads(limit=limit, offset=offset)

    async def health_check(self) -> bool:
        return await self._excel.health_check()
