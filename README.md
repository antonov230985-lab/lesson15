# CSV Processing Bots

Проект переделан под обработку **CSV-файлов** через ботов:

- `main_processor.py` — основной скрипт очистки CSV;
- `telegram_bot_worker.py` — бот Telegram;
- `max_bot_worker.py` — бот MAX;
- `vk_bot_worker.py` — бот группы VK.

## Принцип работы

1. Пользователь отправляет `.csv` файл в бот.
2. Скрипт соответствующего мессенджера скачивает файл локально в `data/`.
3. Вызывается `main_processor.py` (через функцию `process_csv`).
4. Основной скрипт нормализует данные:
   - телефоны -> `79991112233`
   - даты -> `YYYY-MM-DD`
   - суммы -> `float`
   - статусы -> `completed` / `in_progress` / `cancelled`
5. Удаляются полные дубликаты.
6. Готовый CSV отправляется обратно в тот же мессенджер.

## Общий `.env`

Все переменные находятся в одном файле `.env`.
Шаблон — в `.env.example`.

Минимальный запуск:

1. Создать/проверить виртуальное окружение:
   - Windows PowerShell: `python -m venv .venv`
   - активация: `.venv\\Scripts\\Activate.ps1`
2. Установить зависимости: `pip install -r requirements.txt`
3. Скопировать шаблон: `copy .env.example .env`
4. Заполнить токены и идентификаторы в `.env`.
5. Запустить нужный воркер:
   - `python telegram_bot_worker.py`
   - `python max_bot_worker.py`
   - `python vk_bot_worker.py`

## Переменные окружения

- `WORK_DIR` — локальная папка для входных/выходных CSV (по умолчанию `data`)
- `TELEGRAM_BOT_TOKEN`, `TELEGRAM_CHAT_ID`
- `MAX_API_BASE_URL`, `MAX_BOT_TOKEN`, `MAX_CHAT_ID`
- `VK_GROUP_TOKEN`, `VK_GROUP_ID`, `VK_API_VERSION`

## Ручной запуск основного обработчика

```bash
python main_processor.py --input data/input.csv --output data/output.csv
```
