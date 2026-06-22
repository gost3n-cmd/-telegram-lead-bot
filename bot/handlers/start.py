"""Handler: /start (Presentation layer).

Раздел 7: /start, приветственное сообщение, онбординг.
Зависит от: keyboards.
"""

from __future__ import annotations

from aiogram import Router
from aiogram.filters import CommandStart
from aiogram.fsm.context import FSMContext
from aiogram.types import Message

from bot.keyboards.common import get_start_keyboard
from bot.states.email_states import EmailStates

router = Router()


@router.message(CommandStart())
async def handle_start(message: Message, state: FSMContext) -> None:
    """Обработчик команды /start.

    Приветствие + установка FSM в waiting_for_email.
    При повторном /start — сбрасывает FSM и запускает сценарий заново.
    """
    await state.set_state(EmailStates.waiting_for_email)
    await message.answer(
        "👋 Привет!\n\n"
        "Я бот для сбора email-адресов.\n\n"
        "Введите ваш email, чтобы подписаться:\n"
        "(или нажмите кнопку ниже)",
        reply_markup=get_start_keyboard(),
    )
    # Устанавливаем состояние — пользователь должен ввести email
    await message.answer(
        "📧 Пожалуйста, введите ваш email:",
        reply_markup=get_start_keyboard(),
    )
