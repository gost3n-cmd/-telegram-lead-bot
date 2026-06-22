"""GoogleSheetsStorage — secondary output layer for live dashboard.

Best-effort, non-blocking append to Google Sheets via gspread + service account.
All failures are logged and swallowed — never propagates to bot flow.
"""

from __future__ import annotations

import asyncio
import os
from concurrent.futures import ThreadPoolExecutor
from typing import Optional

from loguru import logger

from bot.models.lead import Lead

_GSHEETS_HEADERS = [
    "telegram_id",
    "telegram_username",
    "email",
    "submitted_at",
    "source",
]

_GSHEETS_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="gsheets_io")

_GSHEETS_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


class GoogleSheetsStorage:
    """Appends leads to a Google Sheets worksheet for live visibility.

    Designed for best-effort writes only. Does NOT implement StorageInterface —
    it is a secondary sink, not a source of truth. All errors are caught and
    logged; the caller (DualStorage) never sees exceptions from here.
    """

    def __init__(
        self,
        sheet_id: str,
        credentials_path: str,
        sheet_name: str = "Leads",
    ) -> None:
        self._sheet_id = sheet_id
        self._credentials_path = credentials_path
        self._sheet_name = sheet_name

        self._lock = asyncio.Lock()
        self._executor = _GSHEETS_EXECUTOR

        self._client: object = None
        self._worksheet: object = None
        self._initialized: bool = False

        logger.info(
            f"GoogleSheetsStorage initialized | sheet_id={sheet_id} "
            f"| creds={credentials_path} | sheet={sheet_name}"
        )

    def _run_sync(self, func, *args):
        """Execute a synchronous gspread call in the thread pool."""
        loop = asyncio.get_running_loop()
        return loop.run_in_executor(self._executor, func, *args)

    @staticmethod
    def _authenticate_sync(credentials_path: str, sheet_id: str, sheet_name: str):
        """Synchronous auth + worksheet lookup. Runs in executor."""
        import gspread
        from google.oauth2.service_account import Credentials

        creds = Credentials.from_service_account_file(
            credentials_path, scopes=_GSHEETS_SCOPES
        )
        client = gspread.authorize(creds)
        spreadsheet = client.open_by_key(sheet_id)

        try:
            worksheet = spreadsheet.worksheet(sheet_name)
            logger.debug(f"GS: found existing worksheet '{sheet_name}'")
        except gspread.WorksheetNotFound:
            worksheet = spreadsheet.add_worksheet(title=sheet_name, rows=100, cols=5)
            worksheet.append_row(_GSHEETS_HEADERS)
            logger.info(f"GS: created worksheet '{sheet_name}' with headers")

        return client, worksheet

    async def _ensure_initialized(self) -> None:
        """Lazy init: authenticate and resolve worksheet on first use."""
        if self._initialized:
            return
        async with self._lock:
            if self._initialized:
                return
            if not os.path.exists(self._credentials_path):
                logger.error(
                    f"GS credentials file not found: {self._credentials_path}"
                )
                raise FileNotFoundError(
                    f"Credentials file not found: {self._credentials_path}"
                )
            try:
                self._client, self._worksheet = await self._run_sync(
                    self._authenticate_sync,
                    self._credentials_path,
                    self._sheet_id,
                    self._sheet_name,
                )
                self._initialized = True
                logger.info("GS: authenticated and worksheet ready")
            except Exception as exc:
                logger.error(f"GS initialization failed: {exc}")
                raise

    async def append_lead(self, lead: Lead) -> None:
        """Append one lead row to Google Sheets. Best-effort, never raises."""
        try:
            await self._ensure_initialized()

            row = [
                lead.telegram_id,
                lead.telegram_username or "",
                lead.email,
                lead.submitted_at.isoformat(),
                lead.source,
            ]

            async with self._lock:
                await self._run_sync(self._worksheet.append_row, row)

            logger.debug(f"GS row appended: {lead.email}")
        except FileNotFoundError:
            logger.error(
                f"GS credentials file missing at {self._credentials_path} — skipping"
            )
        except Exception as exc:
            logger.error(f"GS append_lead failed for {lead.email}: {exc}")
