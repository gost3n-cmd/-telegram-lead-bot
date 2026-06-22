"""Единая настройка логирования для всего проекта (Infrastructure layer).

Использует loguru для структурированного логирования.
"""

from __future__ import annotations

import sys

from loguru import logger


def setup_logger(level: str = "INFO") -> None:
    """Настроить loguru: формат, уровень, вывод в stderr."""
    logger.remove()
    logger.add(
        sys.stderr,
        format=(
            "<green>{time:YYYY-MM-DD HH:mm:ss}</green> | "
            "<level>{level: <8}</level> | "
            "<cyan>{name}</cyan>:<cyan>{function}</cyan>:<cyan>{line}</cyan> | "
            "<level>{message}</level>"
        ),
        level=level.upper(),
        colorize=True,
    )
