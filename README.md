# Excel Processing Bots

Проект обрабатывает **Excel-файлы** через ботов:

- `main_processor.py` — основной скрипт очистки Excel;
- `telegram_bot_worker.py` — бот Telegram;
- `max_bot_worker.py` — бот MAX;
- `vk_bot_worker.py` — бот группы VK.

## Принцип работы

1. Пользователь отправляет `.xlsx/.xls` файл в бот.
2. Скрипт мессенджера скачивает файл локально в `data/`.
3. Вызывается `main_processor.py` (через функцию `process_excel_to_workbook`).
4. Основной скрипт выполняет обработку 5 листов:
   - телефоны -> `79991112233`
   - даты -> `YYYY-MM-DD`
   - суммы -> `float`
   - статусы -> `completed` / `in_progress` / `cancelled`
5. Формируется итоговый `.xlsx` с 5 листами.
6. Готовый Excel отправляется обратно в тот же мессенджер.

## Общий `.env`

Все переменные находятся в одном файле `.env`.
Шаблон — в `.env.example`.

Минимальный запуск:

1. Создать/проверить виртуальное окружение:
   - Windows PowerShell: `python -m venv .venv`
   - активация: `.venv\\Scripts\\Activate.ps1`
2. Установить зависимости: `pip install -r requirements.txt`
3. Скопировать шаблон: `copy .env.example .env`
4. Заполнить токены в `.env`.
5. Запустить нужный воркер:
   - `python telegram_bot_worker.py`
   - `python max_bot_worker.py`
   - `python vk_bot_worker.py`

## Имена файлов в ответе

- Telegram: `Ответ в Телеграмм.xlsx`
- MAX: `Ответ в MAX.xlsx`
- VK: `Ответ в ВК.xlsx`

## Переменные окружения

- `WORK_DIR` — локальная папка для временных файлов (по умолчанию `data`)
- `TELEGRAM_BOT_TOKEN`
- `MAX_BOT_TOKEN`
- `VK_GROUP_TOKEN`, `VK_GROUP_ID`, `VK_API_VERSION`

## Важно для VK

Для работы воркера VK в ключе сообщества должны быть включены права:
- сообщения сообщества;
- документы сообщества;
- управление сообществом (для Long Poll API).
