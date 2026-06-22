"""Конфигурация бота через pydantic-settings (Композиция).

Раздел 16 архитектуры: BOT_TOKEN, EXCEL_FILE_PATH, LEADS_SHEET_NAME,
LOG_LEVEL, ENVIRONMENT, GOOGLE_SHEET_ID.
GOOGLE_CREDENTIALS_JSON читается напрямую из os.environ в google_sheets.py.
"""

from __future__ import annotations

from typing import Optional

from pydantic import Field, SecretStr
from pydantic_settings import BaseSettings


class BotConfig(BaseSettings):
    """Конфигурация приложения. Загружается из .env / переменных окружения."""

    bot_token: SecretStr = Field(..., description="Telegram Bot API token")
    excel_file_path: str = Field(
        default="data/ai_content_os.xlsx",
        description="Путь к общему Excel-файлу",
    )
    leads_sheet_name: str = Field(
        default="Leads",
        description="Имя листа бота в Excel-книге (не должно совпадать с OWNED_SHEETS генератора)",
    )
    log_level: str = Field(default="INFO", description="Уровень логирования")
    environment: str = Field(default="production", description="production или development")

    google_sheet_id: Optional[str] = Field(
        default=None,
        description="ID Google Sheets таблицы для live-дашборда",
    )
    # GOOGLE_CREDENTIALS_JSON читается напрямую из os.environ в google_sheets.py

    model_config = {"env_file": ".env", "env_file_encoding": "utf-8"}
config = BotConfig()