"""FSM-состояния для сбора email (Presentation layer).

Раздел 13.1: для MVP достаточно одного состояния.
"""

from aiogram.fsm.state import StatesGroup, State


class EmailStates(StatesGroup):
    """Состояния конечного автомата для сценария сбора email."""

    waiting_for_email = State()
