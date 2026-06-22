"""Handler: сбор email через FSM (Presentation layer).

Раздел 7: принимает текстовое сообщение, переводит FSM в нужное состояние,
вызывает lead_service, отправляет пользователю результат.

DI: lead_service передаётся через фабрику create_email_router (closure),
что гарантирует, что хендлер не знает о конкретной реализации хранилища.

Зависит от: states, services.lead_service, keyboards.
НЕ импортирует storage/excel_storage.py напрямую.
"""

from __future__ import annotations

from aiogram import Router, F
from aiogram.fsm.context import FSMContext
from aiogram.types import Message

from bot.services.lead_service import LeadService, ServiceErrorType
from bot.states.email_states import EmailStates
from bot.keyboards.common import get_cancel_keyboard


def create_email_router(lead_service: LeadService) -> Router:
    """Фабрика роутера с внедрённым lead_service через closure.

    Это паттерн DI: сервис передаётся из main.py при сборке приложения.
    """

    router = Router()

    @router.message(EmailStates.waiting_for_email, F.text)
    async def handle_email_input(message: Message, state: FSMContext) -> None:
        """Обработчик текстового ввода в состоянии waiting_for_email.

        Принимает email, передаёт в lead_service, отправляет результат пользователю.
        """
        # Кнопка отмены
        if message.text == "❌ Отмена":
            await state.clear()
            await message.answer(
                "❌ Отменено.\nНапишите /start чтобы начать заново.",
                reply_markup=None,
            )
            return

        # Делегируем сервису
        result = await lead_service.process_email(
            raw_email=message.text,
            telegram_id=message.from_user.id,
            telegram_username=message.from_user.username,
        )

        if result.success:
            # Успех — сбрасываем FSM
            await state.clear()
            await message.answer(
                "✅ Email сохранён!\n\n"
                "Спасибо за подписку. Вы будете получать наши обновления.",
                reply_markup=None,
            )
        else:
            # Ошибка — FSM остаётся в waiting_for_email
            error_messages = {
                ServiceErrorType.VALIDATION_FAILED: (
                    f"❌ {result.error_message}\n\n"
                    "Пожалуйста, введите корректный email:"
                ),
                ServiceErrorType.DUPLICATE: (
                    f"⚠️ {result.error_message}\n\n"
                    "Введите другой email или нажмите ❌ Отмена:"
                ),
                ServiceErrorType.STORAGE_UNAVAILABLE: (
                    "🔧 Что-то пошло не так — попробуйте позже."
                ),
                ServiceErrorType.UNKNOWN_ERROR: (
                    "🔧 Что-то пошло не так — попробуйте позже."
                ),
            }
            text = error_messages.get(
                result.error_type,
                "🔧 Произошла ошибка — попробуйте позже.",
            )
            await message.answer(text, reply_markup=get_cancel_keyboard())

    @router.message(EmailStates.waiting_for_email)
    async def handle_non_text_input(message: Message, state: FSMContext) -> None:
        """Обработчик нетекстового ввода в состоянии waiting_for_email."""
        await message.answer(
            "⚠️ Пожалуйста, введите email текстом.\n\n"
            "Или нажмите ❌ Отмена:",
            reply_markup=get_cancel_keyboard(),
        )

    @router.message()
    async def handle_fallback(message: Message) -> None:
        """Fallback-хендлер для сообщений вне FSM.

        Раздел 13.3: подсказывает /start.
        """
        await message.answer(
            "👋 Чтобы начать, отправьте команду /start",
        )

    return router
