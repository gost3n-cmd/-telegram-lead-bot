"""Точка входа бота — сборка зависимостей, DI, запуск polling (Композиция).

Раздел 5/7/9: единственное место, где упоминается конкретная реализация хранилища.
Создаёт Bot/Dispatcher, ExcelStorage, регистрирует роутеры и middleware.
"""

from __future__ import annotations

import asyncio

from aiogram import Bot, Dispatcher
from aiogram.client.default import DefaultBotProperties
from aiogram.fsm.storage.memory import MemoryStorage
from loguru import logger

from bot.config import BotConfig
from bot.handlers.start import router as start_router
from bot.handlers.email_collection import create_email_router
from bot.middlewares.logging_middleware import LoggingMiddleware
from bot.middlewares.throttling_middleware import ThrottlingMiddleware
from bot.services.lead_service import LeadService
from bot.storage.excel_storage import ExcelStorage
from bot.utils.logger import setup_logger


def _build_storage(config: BotConfig):
    """Build storage stack: Excel (primary) + optional Google Sheets (secondary)."""
    excel = ExcelStorage(
        file_path=config.excel_file_path,
        sheet_name=config.leads_sheet_name,
    )

    if config.google_sheet_id and config.google_credentials_path:
        from bot.storage.google_sheets import GoogleSheetsStorage
        from bot.storage.dual_storage import DualStorage

        gsheets = GoogleSheetsStorage(
            sheet_id=config.google_sheet_id,
            credentials_path=config.google_credentials_path,
        )
        logger.info("Google Sheets secondary storage enabled")
        return DualStorage(excel_storage=excel, gsheets_storage=gsheets)

    logger.info("Google Sheets not configured — Excel-only mode")
    return excel


def build_application(config: BotConfig) -> tuple[Bot, Dispatcher]:
    """Собрать приложение: Bot, Dispatcher, DI-зависимости.

    Единственная точка, где импортируются конкретные реализации хранилища.
    """
    # Логирование
    setup_logger(level=config.log_level)

    # Bot
    bot = Bot(
        token=config.bot_token.get_secret_value(),
        default=DefaultBotProperties(parse_mode="HTML"),
    )

    # Dispatcher с MemoryStorage для FSM
    dp = Dispatcher(storage=MemoryStorage())

    # Infrastructure: хранилище (Excel + опционально Google Sheets)
    storage = _build_storage(config)
    logger.info(f"STORAGE TYPE: {type(storage).__name__}")
    logger.info(f"Excel path: {config.excel_file_path}")
    logger.info(f"Storage class: {storage.__class__.__name__}")

    # Application: сервис с внедрённой зависимостью (StorageInterface)
    lead_service = LeadService(storage=storage)

    # Presentation: регистрация роутеров
    dp.include_router(start_router)

    # email_router получает lead_service через фабрику — DI через closure
    email_router = create_email_router(lead_service)
    dp.include_router(email_router)

    # Middlewares
    dp.update.middleware(LoggingMiddleware())
    dp.message.middleware(ThrottlingMiddleware(max_messages=5, window_seconds=10))

    logger.info("Application built successfully")
    return bot, dp


async def main() -> None:
    """Главная асинхронная функция."""
    # Загрузка конфигурации
    config = BotConfig()

    bot, dp = build_application(config)

    logger.info(f"Starting bot in {config.environment} mode")
    logger.info(f"Excel file: {config.excel_file_path}")
    logger.info(f"Leads sheet: {config.leads_sheet_name}")

    # Запуск polling
    try:
        await dp.start_polling(bot)
    finally:
        await bot.session.close()
        logger.info("Bot stopped")


def run() -> None:
    """Синхронная точка входа для `python -m bot.main`."""
    asyncio.run(main())


if __name__ == "__main__":
    run()
