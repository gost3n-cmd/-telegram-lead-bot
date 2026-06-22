"""Валидация email — Domain-логика (бизнес-правила формата).

Проверка синтаксиса email, нормализация (нижний регистр, обрезка пробелов).
Не обращается к диску и не знает про Excel.
"""

from __future__ import annotations

import re
from dataclasses import dataclass


# Базовый паттерн для валидации email (не RFC-полный, но покрывает 99% случаев)
_EMAIL_PATTERN = re.compile(
    r"^[a-zA-Z0-9!#$%&'*+/=?^_`{|}~-]+"
    r"(?:\.[a-zA-Z0-9!#$%&'*+/=?^_`{|}~-]+)*"
    r"@"
    r"(?:[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?\.)+"
    r"[a-zA-Z]{2,63}$"
)

# Максимальная длина email по RFC 5321
_MAX_EMAIL_LENGTH = 254


@dataclass(frozen=True)
class ValidationResult:
    """Результат валидации email."""

    is_valid: bool
    email: str | None = None
    error: str | None = None


def validate_email(raw_email: str) -> ValidationResult:
    """Валидировать и нормализовать email.

    1. Обрезка пробелов по краям
    2. Нижний регистр
    3. Проверка длины
    4. Проверка синтаксиса через regex

    Возвращает ValidationResult с нормализованным email или ошибкой.
    """
    if not raw_email or not isinstance(raw_email, str):
        return ValidationResult(is_valid=False, error="Email не может быть пустым")

    normalized = raw_email.strip().lower()

    if not normalized:
        return ValidationResult(is_valid=False, error="Email не может быть пустым")

    if len(normalized) > _MAX_EMAIL_LENGTH:
        return ValidationResult(
            is_valid=False,
            error=f"Email слишком длинный (максимум {_MAX_EMAIL_LENGTH} символов)",
        )

    if " " in normalized:
        return ValidationResult(is_valid=False, error="Email не должен содержать пробелы")

    if not _EMAIL_PATTERN.match(normalized):
        return ValidationResult(is_valid=False, error="Некорректный формат email")

    return ValidationResult(is_valid=True, email=normalized)
