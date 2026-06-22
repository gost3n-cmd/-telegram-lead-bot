"""LeadService — оркестрация сценария "принять email" (Application layer).

Единственная точка, где "встречаются" Validation и Storage:
сырой текст → валидирует → проверяет дубликат → сохраняет → возвращает результат.

Раздел 15.2: перехватывает Storage-исключения, преобразует в ServiceResult.
НЕ пробрасывает Storage-исключения наружу.
"""

from __future__ import annotations

import enum
from dataclasses import dataclass
from typing import Optional

from loguru import logger

from bot.models.lead import Lead
from bot.storage.base import StorageInterface
from bot.storage.exceptions import StorageUnavailableError
from bot.validation.email_validator import validate_email, ValidationResult


class ServiceErrorType(enum.Enum):
    """Типы ошибок на уровне сервиса."""

    VALIDATION_FAILED = "validation_failed"
    DUPLICATE = "duplicate"
    STORAGE_UNAVAILABLE = "storage_unavailable"
    UNKNOWN_ERROR = "unknown_error"


@dataclass
class ServiceResult:
    """Результат работы lead_service — возвращается хендлеру."""

    success: bool
    error_type: Optional[ServiceErrorType] = None
    error_message: Optional[str] = None
    lead: Optional[Lead] = None


class LeadService:
    """Сервис оркестрации сохранения лида.

    Знает про Validation и Storage (через интерфейс).
    Не знает про aiogram и не знает про openpyxl.
    """

    def __init__(self, storage: StorageInterface) -> None:
        self._storage = storage

    async def process_email(
        self,
        raw_email: str,
        telegram_id: int,
        telegram_username: Optional[str] = None,
    ) -> ServiceResult:
        """Полный цикл: валидация → проверка дубликата → сохранение.

        Args:
            raw_email: Сырой email от пользователя.
            telegram_id: Telegram ID пользователя.
            telegram_username: Username (может быть None).

        Returns:
            ServiceResult с результатом операции.
        """
        # 1. Валидация
        validation = validate_email(raw_email)
        if not validation.is_valid:
            logger.debug(f"Validation failed: tg_id={telegram_id}, error={validation.error}")
            return ServiceResult(
                success=False,
                error_type=ServiceErrorType.VALIDATION_FAILED,
                error_message=validation.error,
            )

        email = validation.email

        # 2. Сохранение с атомарной проверкой дубликата внутри блокировки
        lead = Lead(
            telegram_id=telegram_id,
            telegram_username=telegram_username,
            email=email,
        )

        try:
            save_result = await self._storage.save_lead(lead)
            if not save_result.success:
                if save_result.is_duplicate:
                    logger.info(f"Duplicate email: tg_id={telegram_id}, email={email}")
                    return ServiceResult(
                        success=False,
                        error_type=ServiceErrorType.DUPLICATE,
                        error_message=save_result.error or "Этот email уже зарегистрирован",
                    )
                logger.error(f"Save failed: email={email}, error={save_result.error}")
                return ServiceResult(
                    success=False,
                    error_type=ServiceErrorType.STORAGE_UNAVAILABLE,
                    error_message=save_result.error or "Ошибка сохранения",
                )
        except StorageUnavailableError as exc:
            logger.error(f"Storage unavailable during save: {exc.message}")
            return ServiceResult(
                success=False,
                error_type=ServiceErrorType.STORAGE_UNAVAILABLE,
                error_message="Хранилище недоступно",
            )
        except Exception as exc:
            logger.error(f"Unexpected error during save: {exc}")
            return ServiceResult(
                success=False,
                error_type=ServiceErrorType.UNKNOWN_ERROR,
                error_message="Произошла непредвиденная ошибка",
            )

        logger.info(f"Lead processed successfully: email={email}, tg_id={telegram_id}")
        return ServiceResult(success=True, lead=lead)
