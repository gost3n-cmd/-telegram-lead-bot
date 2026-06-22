# TODO — Задачи для будущих итераций

## Критические
- [ ] Обновить `generator/build_system.py` с Sheet Ownership (раздел 4 архитектуры)
  - Добавить `OWNED_SHEETS` константу (11 имён)
  - Заменить `wb = Workbook()` на `load_workbook()` при существующем файле
  - Удалять только owned sheets перед пересозданием
  - Переупорядочивать owned sheets в начало (косметика)
  - После внедрения — удалить оригинальный `build_system.py` из корня проекта

## Средний приоритет
- [ ] Создать `README.md` с инструкциями по запуску
- [ ] Реальный smoke-тест бота с валидным Telegram-токеном
- [ ] Интеграция `email-validator` библиотеки при необходимости строгой RFC-валидации (D011)
- [ ] Webhook-режим деплоя (раздел 12 — открытый вопрос)
- [ ] Стилизация листа Leads под общий дизайн книги (раздел 12)

## Низкий приоритет
- [ ] Скрипт миграции `scripts/migrate_excel_to_sqlite.py` (будущий этап)
- [ ] Админ-команды: `/stats`, `/export` (используют `count_leads`, `list_leads`)
- [ ] File-level lock (`filelock`/`portalocker`) при запуске бота в нескольких процессах (раздел 11)
- [ ] Команда `/cancel` как альтернатива кнопке "Отмена"
- [ ] Хранение персональных данных и право на удаление (раздел 12 — открытый вопрос)

## Завершено (Итерация 1, 2026-06-22)
- [x] Структура проекта по разделу 6
- [x] Все слои архитектуры (models → storage contracts → exceptions → lock → validation → services → handlers → main)
- [x] Тесты для validator и storage (31 тест, все проходят)
- [x] Документация: `docs/decisions.md`, `docs/implementation_progress.md`, `project_status.md`
