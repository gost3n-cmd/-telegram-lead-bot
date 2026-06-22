"""Клавиатуры (Presentation layer).

Кнопки для взаимодействия с пользователем.
"""

from __future__ import annotations

from aiogram.types import ReplyKeyboardMarkup, KeyboardButton


def get_start_keyboard() -> ReplyKeyboardMarkup:
    """Клавиатура после /start — запрос email."""
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="📧 Ввести email")],
        ],
        resize_keyboard=True,
    )


def get_cancel_keyboard() -> ReplyKeyboardMarkup:
    """Клавиатура с кнопкой отмены (в состоянии waiting_for_email)."""
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="❌ Отмена")],
        ],
        resize_keyboard=True,
    )
