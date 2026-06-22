"""Конфигурация бота через pydantic-settings (Композиция).

Раздел 16 архитектуры: BOT_TOKEN, EXCEL_FILE_PATH, LEADS_SHEET_NAME,
LOG_LEVEL, ENVIRONMENT.
"""

from __future__ import annotations

from pydantic import Field, SecretStr
from pydantic_settings import BaseSettings


class BotConfig(BaseSettings):
    """Конфигурация приложения. Загружается из .env / переменных окружения."""

    bot_token: SecretStr = Field(..., description="8973436711:AAHrXw2J88KEHhBngTIsTM7jWv_JkLA1uRY")
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

    model_config = {"env_file": ".env", "env_file_encoding": "utf-8"}
