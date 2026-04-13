# Lesson 13 — Data Cleaning Homework

Небольшой учебный проект по очистке данных из Excel (`homework_lesson13.xlsx`) с помощью Python + pandas.

## Что делает код

Скрипт `solve_lesson13.py`:
- читает 5 листов с "грязными" данными;
- нормализует телефоны, даты, суммы и статусы;
- удаляет служебные/пустые строки и дубли;
- агрегирует заказы (несколько строк -> одна строка);
- сохраняет чистые результаты в отдельные и общий Excel-файлы.

## Ключевые форматы на выходе

- телефон: `79991112233`
- дата: `2025-01-15`
- сумма: `1250.0`
- статус: `completed` / `in_progress` / `cancelled`

## Как запустить

```bash
python solve_lesson13.py
```

## Что создается после запуска

- `task1_phones_dates_clean.xlsx`
- `task2_amounts_statuses_clean.xlsx`
- `task3_garbage_duplicates_clean.xlsx`
- `task4_orders_aggregated_clean.xlsx`
- `task5_full_pipeline_clean.xlsx`
- `lesson13_all_tasks_clean.xlsx` (общий файл с 5 листами)

## Зависимости

- Python 3.10+
- `pandas`
- `openpyxl`

Установка:

```bash
pip install pandas openpyxl
```
