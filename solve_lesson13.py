"""
Этот файл решает все 5 заданий из домашней работы по уроку 13.

Ниже логика написана максимально "человеческим" языком:
- что берется из исходного Excel,
- как именно нормализуются данные,
- куда и в каком формате сохраняется результат.
"""

# Модуль re нужен для регулярных выражений:
# это удобный способ "чистить" строки и искать шаблоны (например, оставить только цифры).
import re

# Path из pathlib нужен, чтобы работать с путями к файлам удобно и безопасно.
from pathlib import Path

# pandas - главная библиотека здесь: читает Excel, чистит таблицы, сохраняет обратно в Excel.
import pandas as pd


# Словарь для русских месяцев.
# Зачем: pandas лучше понимает английские сокращения месяцев.
# Поэтому мы заранее заменяем русский месяц на английский эквивалент:
# "мар" -> "mar", "июн" -> "jun" и т.д.
RU_MONTHS = {
    "янв": "jan",
    "фев": "feb",
    "мар": "mar",
    "апр": "apr",
    "мая": "may",
    "май": "may",
    "июн": "jun",
    "июл": "jul",
    "авг": "aug",
    "сен": "sep",
    "сент": "sep",
    "окт": "oct",
    "ноя": "nov",
    "дек": "dec",
}

# Словарь для приведения "грязных" статусов к финальным значениям по заданию.
# На выходе везде должно быть только одно из трех:
# completed / in_progress / cancelled.
STATUS_MAP = {
    "завершен": "completed",
    "завершён": "completed",
    "закрыт": "completed",
    "в работе": "in_progress",
    "отменен": "cancelled",
    "отменён": "cancelled",
}


def norm_phone(value: object) -> str | None:
    """
    Нормализует телефон к формату 11 цифр, начинающихся с 7.

    Что функция получает:
    - любое значение из ячейки Excel (строка, число, пустое).

    Что делает пошагово:
    1) Если ячейка пустая -> возвращает None (это "пустое значение" в Python).
    2) Превращает значение в строку и убирает пробелы по краям.
    3) Если есть хвост ".0" (часто после Excel), убирает его.
    4) Удаляет вообще все, кроме цифр.
    5) Если осталось 10 цифр -> добавляет "7" в начало.
    6) Если 11 цифр и начинается на "8" -> заменяет первую цифру на "7".
    7) Если получилось корректно: 11 цифр и первая цифра "7" -> возвращает телефон.
    8) Иначе возвращает None (значит номер не удалось привести к стандарту).

    Пример:
    "+7 (999) 111-22-33" -> "79991112233"
    """
    if pd.isna(value):          
        return None

    # Приводим к строке, чтобы одинаково обработать и текст, и числа.
    text = str(value).strip()

    # Удаляем типичный excel-артефакт: "79990001122.0".
    if text.endswith(".0"):
        text = text[:-2]

    # \D означает "не цифра". Заменяем все не-цифры на пустую строку.
    digits = re.sub(r"\D", "", text)

    # Если номер без кода страны (10 цифр), добавляем "7".
    if len(digits) == 10:
        digits = "7" + digits
    # Если начинается с "8", меняем на "7" (российский формат).
    elif len(digits) == 11 and digits.startswith("8"):
        digits = "7" + digits[1:]

    # Возвращаем только валидный итог.
    if len(digits) == 11 and digits.startswith("7"):
        return digits
    return None


def norm_date(value: object) -> str | None:
    """
    Нормализует дату к формату YYYY-MM-DD.

    Что функция получает:
    - значение даты в любом "ломаном" виде:
      "15.01.2025", "2025/02/20", "20 мар 2025", "15 июня 2025", "01.03.25" и т.д.

    Как обрабатывает:
    1) Пустое значение -> None.
    2) Приводит строку к нижнему регистру.
    3) Делает все разделители одинаковыми (пробел):
       "-", "/", "." -> " ".
    4) Русские месяцы заменяет на английские (через RU_MONTHS).
    5) Сжимает повторные пробелы до одного.
    6) Пытается распарсить в дату:
       - сначала строгий формат "год месяц день" (чтобы убрать предупреждения),
       - потом обычный парсинг с dayfirst=True и dayfirst=False.
    7) Если удалось распознать дату -> возвращает в формате "YYYY-MM-DD".
       Если не удалось -> None.
    """
    if pd.isna(value):
        return None

    text = str(value).strip().lower()

    # Унифицируем разделители, чтобы дальше парсеру было проще.
    text = text.replace("-", " ").replace("/", " ").replace(".", " ")

    # Меняем русские названия месяцев на английские.
    for ru, en in RU_MONTHS.items():
        text = re.sub(rf"\b{ru}\w*\b", en, text)

    # Удаляем лишние пробелы.
    text = re.sub(r"\s+", " ", text).strip()

    # Если строка выглядит как "2025 03 10" - используем точный формат.
    if re.match(r"^\d{4}\s+\d{1,2}\s+\d{1,2}$", text):
        dt = pd.to_datetime(text, errors="coerce", format="%Y %m %d")
        if pd.notna(dt):
            return dt.strftime("%Y-%m-%d")

    # Для других случаев делаем две попытки:
    # сначала "день-месяц-год", потом "месяц-день-год".
    for dayfirst in (True, False):
        dt = pd.to_datetime(text, errors="coerce", dayfirst=dayfirst)
        if pd.notna(dt):
            return dt.strftime("%Y-%m-%d")
    return None


def norm_amount(value: object) -> float | None:
    """
    Нормализует сумму к типу float.

    Что функция умеет исправлять:
    - "1 250" -> 1250.0
    - "1250 ₽" -> 1250.0
    - "1850,00" -> 1850.0
    - "1,850.00" -> 1850.0
    - "3500р" -> 3500.0
    - "нет данных" -> None

    Шаги:
    1) Пустая ячейка -> None.
    2) Убираем валютные символы и пробелы.
    3) Отсекаем явно нечисловые значения (например "нет данных").
    4) Разбираемся с запятыми/точками:
       - если есть и запятая, и точка -> определяем, кто из них дробная часть;
       - если только запятая -> решаем, это десятичный разделитель или разделитель тысяч.
    5) Пробуем float(...). Если не получилось -> None.
    """
    if pd.isna(value):
        return None

    text = str(value).strip().lower()
    text = text.replace("₽", "").replace("р", "").replace(" ", "")

    # Значения, которые считаем "не суммой".
    if text in {"", "нетданных", "nan"}:
        return None

    # Одновременно запятая и точка:
    # пример "1,850.00" или "1.850,00".
    if "," in text and "." in text:
        # Последний символ-разделитель обычно дробная часть.
        if text.rfind(",") > text.rfind("."):
            # Формат типа "1.850,00": точка - тысячи, запятая - дробь.
            text = text.replace(".", "").replace(",", ".")
        else:
            # Формат типа "1,850.00": запятая - тысячи, точка - дробь.
            text = text.replace(",", "")
    elif "," in text:
        # Только запятая: это может быть либо дробь, либо разделитель тысяч.
        parts = text.split(",")
        # Если справа 1-2 цифры, скорее всего это дробная часть.
        if len(parts) == 2 and len(parts[1]) in {1, 2}:
            text = text.replace(",", ".")
        else:
            # Иначе считаем разделителем тысяч и удаляем.
            text = text.replace(",", "")

    try:
        return float(text)
    except ValueError:
        return None


def norm_status(value: object) -> str | None:
    """
    Нормализует статус по словарю STATUS_MAP.

    Пример:
    - "Завершён" -> "completed"
    - "В работе" -> "in_progress"
    - "отменен" -> "cancelled"
    - неизвестный текст -> None
    """
    if pd.isna(value):
        return None
    key = str(value).strip().lower()
    return STATUS_MAP.get(key)


def solve_task_1(source: Path, out_path: Path) -> pd.DataFrame:
    """
    Решение задания 1: телефоны и даты.

    Откуда берем данные:
    - из файла source,
    - лист "1_Телефоны_и_даты",
    - header=2, потому что первые 2 строки на листе - это текст инструкции.

    Что делаем:
    - оставляем ID и Клиент для идентификации,
    - к колонке "Телефон (исходный)" применяем norm_phone(),
    - к колонке "Дата (исходная)" применяем norm_date(),
    - сохраняем в out_path.

    Что возвращаем:
    - готовый DataFrame, чтобы его можно было использовать дальше
      (например, записать в общий файл со всеми заданиями).
    """
    df = pd.read_excel(source, sheet_name="1_Телефоны_и_даты", header=2)
    result = df[["ID", "Клиент"]].copy()
    result["phone"] = df["Телефон (исходный)"].apply(norm_phone)
    result["date"] = df["Дата (исходная)"].apply(norm_date)
    result.to_excel(out_path, index=False)
    return result


def solve_task_2(source: Path, out_path: Path) -> pd.DataFrame:
    """
    Решение задания 2: суммы и статусы.

    Откуда данные:
    - лист "2_Суммы_и_статусы".

    Что делаем:
    - оставляем ID и Услуга,
    - чистим сумму функцией norm_amount(),
    - чистим статус функцией norm_status(),
    - сохраняем отдельный чистый файл.
    """
    df = pd.read_excel(source, sheet_name="2_Суммы_и_статусы", header=2)
    result = df[["ID", "Услуга"]].copy()
    result["amount"] = df["Сумма (исходная)"].apply(norm_amount)
    result["status"] = df["Статус (исходный)"].apply(norm_status)
    result.to_excel(out_path, index=False)
    return result


def solve_task_3(source: Path, out_path: Path) -> pd.DataFrame:
    """
    Решение задания 3: мусор и дубли.

    Здесь важная идея:
    - сначала вычищаем явно нерабочие строки (служебные, пустые),
    - потом отмечаем точные дубли,
    - для не-дублей помечаем, является ли визит клиента первым или повторным.

    Почему так:
    - "повторный визит" может быть нормальным событием (не ошибка),
      если дата/услуга отличается от предыдущей записи.
    - "точный дубль" (все ключевые поля совпали) чаще всего техническая копия.
      Такие строки не удаляем сразу, а помечаем как "дубль" и очищаем ID.
    """
    df = pd.read_excel(source, sheet_name="3_Мусор_и_дубли", header=2)

    # Копия исходной таблицы, чтобы не портить оригинал.
    work = df.copy()

    # Приводим и подготавливаем поля для удобной фильтрации и сравнения.
    work["id_num"] = pd.to_numeric(work["#"], errors="coerce")
    work["amount_num"] = work["Сумма"].apply(norm_amount)
    work["phone_norm"] = work["Телефон"].apply(norm_phone)
    work["date_norm"] = work["Дата визита"].apply(norm_date)

    # Выкидываем строки, где нет признаков нормальной записи визита:
    # нет числового ID, нет даты, нет клиента или услуги.
    clean = work[
        work["id_num"].notna()
        & work["date_norm"].notna()
        & work["Клиент"].notna()
        & work["Услуга"].notna()
    ].copy()

    # Определяем "точный дубль":
    # если полностью совпадают клиент, телефон, дата, услуга и сумма.
    exact_key = ["Клиент", "phone_norm", "date_norm", "Услуга", "amount_num"]
    clean["is_exact_duplicate"] = clean.duplicated(subset=exact_key, keep="first")

    # Номер визита считаем только для уникальных записей (не дублей).
    uniq = clean[~clean["is_exact_duplicate"]].copy()
    uniq["visit_number_for_client"] = (
        uniq.sort_values("date_norm").groupby(["Клиент", "phone_norm"]).cumcount() + 1
    )
    clean = clean.merge(
        uniq[exact_key + ["visit_number_for_client"]],
        on=exact_key,
        how="left",
    )

    # Маркировка строк:
    # - дубль -> "дубль"
    # - не дубль -> "первичный визит" / "повторный визит"
    clean["duplicate_note"] = clean.apply(
        lambda row: (
            "дубль"
            if row["is_exact_duplicate"]
            else ("повторный визит" if row["visit_number_for_client"] > 1 else "первичный визит")
        ),
        axis=1,
    )

    # Для дублей очищаем ID, чтобы их можно было удалить вручную позже.
    clean["id_out"] = clean["#"]
    clean.loc[clean["is_exact_duplicate"], "id_out"] = None

    # Формируем итоговые колонки и имена в более удобном стиле.
    result = clean[
        ["id_out", "date_norm", "Клиент", "phone_norm", "Услуга", "amount_num", "duplicate_note"]
    ].rename(
        columns={
            "id_out": "#",
            "date_norm": "date",
            "phone_norm": "phone",
            "amount_num": "amount",
        }
    )

    result.to_excel(out_path, index=False)
    return result


def solve_task_4(source: Path, out_path: Path) -> pd.DataFrame:
    """
    Решение задания 4: агрегировать заказ в 1 строку.

    Логика листа:
    - у каждого order_id есть строка типа "Клиент" (в ней дата/клиент/телефон),
    - и несколько строк типа "Работа"/"Товар" (позиции заказа).

    Что считаем:
    - total_amount = сумма (Кол-во * Цена) по всем позициям заказа,
    - items_count = сколько позиций "Работа"/"Товар" в заказе.
    """
    df = pd.read_excel(source, sheet_name="4_Логика_заказов", header=2)

    # Преобразуем числовые поля, где возможно.
    df["order_id"] = pd.to_numeric(df["order_id"], errors="coerce")
    df["price_num"] = pd.to_numeric(df["Цена"], errors="coerce")
    df["qty_num"] = pd.to_numeric(df["Кол-во"], errors="coerce")

    rows: list[dict] = []

    # Проходим по каждому заказу отдельно.
    for order_id, part in df.groupby("order_id", dropna=True):
        # Данные клиента лежат в строке "Клиент".
        client_row = part[part["Тип строки"] == "Клиент"].head(1)

        date = norm_date(client_row["Дата"].iloc[0]) if not client_row.empty else None
        client = client_row["Клиент"].iloc[0] if not client_row.empty else None
        phone = norm_phone(client_row["Телефон"].iloc[0]) if not client_row.empty else None

        # Позиции заказа: работы и товары.
        item_rows = part[part["Тип строки"].isin(["Работа", "Товар"])].copy()
        item_rows["line_total"] = item_rows["qty_num"].fillna(0) * item_rows["price_num"].fillna(0)

        total_amount = float(item_rows["line_total"].sum())
        items_count = int(item_rows.shape[0])

        rows.append(
            {
                "order_id": int(order_id),
                "date": date,
                "client": client,
                "phone": phone,
                "total_amount": total_amount,
                "items_count": items_count,
            }
        )

    result = pd.DataFrame(rows).sort_values("order_id")
    result.to_excel(out_path, index=False)
    return result


def solve_task_5(source: Path, out_path: Path) -> pd.DataFrame:
    """
    Решение задания 5: комплексный pipeline.

    Здесь совмещаются все подходы:
    - удаление мусора и служебных строк,
    - нормализация даты/телефона/суммы/статуса,
    - базовая чистка текстовых полей,
    - удаление полных дублей.
    """
    df = pd.read_excel(source, sheet_name="5_Комплексное", header=2)
    work = df.copy()

    # Оставляем только строки с нормальным номером записи.
    # Служебные строки обычно не имеют числового значения в колонке "#".
    work["id_num"] = pd.to_numeric(work["#"], errors="coerce")
    work = work[work["id_num"].notna()].copy()

    # Нормализация всех ключевых полей.
    work["date"] = work["Дата"].apply(norm_date)
    work["client"] = work["Клиент"].astype(str).str.strip().str.title()
    work["phone"] = work["Телефон"].apply(norm_phone)
    work["vin"] = work["VIN"].astype(str).str.strip()
    work["service"] = work["Услуга"].astype(str).str.strip()
    work["amount"] = work["Сумма"].apply(norm_amount)
    work["status"] = work["Статус"].apply(norm_status)

    # После astype(str) пропуски превращаются в "nan"/"None".
    # Возвращаем их обратно в пустые значения.
    work.loc[work["vin"].isin(["nan", "None"]), "vin"] = None
    work.loc[work["service"].isin(["nan", "None"]), "service"] = None

    # Оставляем валидные строки по ключевым полям.
    # Сумма и статус могут отсутствовать: такие строки сохраняем для ручной проверки.
    clean = work[
        work["date"].notna()
        & work["client"].notna()
        & work["phone"].notna()
        & work["service"].notna()
    ].copy()

    # Удаляем полные дубли.
    clean = clean.drop_duplicates(
        subset=["date", "client", "phone", "vin", "service", "amount", "status"],
        keep="first",
    )

    result = clean[["#", "date", "client", "phone", "vin", "service", "amount", "status"]]
    result.to_excel(out_path, index=False)
    return result


def save_all_tasks_to_one_workbook(results: dict[str, pd.DataFrame], out_path: Path) -> None:
    """
    Сохраняет все готовые таблицы в один общий Excel-файл.

    Что получает:
    - results: словарь, где ключ = имя листа, значение = DataFrame.
    - out_path: путь к итоговому excel-файлу.

    Что делает:
    - открывает ExcelWriter,
    - записывает каждый DataFrame на отдельный лист,
    - сохраняет один общий файл.
    """
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for sheet_name, dataframe in results.items():
            dataframe.to_excel(writer, sheet_name=sheet_name, index=False)


def main() -> None:
    """
    Главная функция запуска всего скрипта.

    Порядок действий:
    1) Указываем исходный файл.
    2) Запускаем 5 решателей (по одному на каждое задание),
       каждый одновременно сохраняет свой отдельный .xlsx
       и возвращает DataFrame в память.
    3) Из этих 5 DataFrame создаем один общий Excel с 5 листами.
    4) Печатаем понятное сообщение об успешном завершении.
    """
    source = Path("homework_lesson13.xlsx")

    # Запускаем каждое задание и сохраняем отдельные файлы.
    task1_df = solve_task_1(source, Path("task1_phones_dates_clean.xlsx"))
    task2_df = solve_task_2(source, Path("task2_amounts_statuses_clean.xlsx"))
    task3_df = solve_task_3(source, Path("task3_garbage_duplicates_clean.xlsx"))
    task4_df = solve_task_4(source, Path("task4_orders_aggregated_clean.xlsx"))
    task5_df = solve_task_5(source, Path("task5_full_pipeline_clean.xlsx"))

    # Готовим словарь для общего файла: имя листа -> данные.
    all_results = {
        "task1_phones_dates": task1_df,
        "task2_amounts_statuses": task2_df,
        "task3_garbage_duplicates": task3_df,
        "task4_orders_aggregated": task4_df,
        "task5_full_pipeline": task5_df,
    }

    # Сохраняем все вместе в один workbook (по просьбе пользователя).
    save_all_tasks_to_one_workbook(all_results, Path("lesson13_all_tasks_clean.xlsx"))

    print(
        "Done: saved 5 separate files and 1 combined workbook "
        "lesson13_all_tasks_clean.xlsx"
    )


# Эта проверка нужна, чтобы main() выполнялась только при прямом запуске файла.
# Если файл импортируют в другой скрипт как модуль, main() автоматически не запустится.
if __name__ == "__main__":
    main()


# =============================================================================
# FAQ ДЛЯ НОВИЧКА (ПРОСТЫМИ СЛОВАМИ)
# =============================================================================
#
# 1) Что такое DataFrame?
#    DataFrame (из pandas) - это таблица в памяти Python.
#    По сути, "умный Excel-лист": есть строки, колонки, фильтры, сортировка, расчеты.
#
# 2) Что делает pd.read_excel(...)?
#    Эта команда открывает Excel-файл и читает выбранный лист в DataFrame.
#    Пример:
#        df = pd.read_excel("file.xlsx", sheet_name="Лист1")
#    После этого переменная df содержит таблицу, с которой можно работать в коде.
#
# 3) Почему в read_excel стоит header=2?
#    Потому что в ваших листах первые 2 строки - это инструкции/текст,
#    а настоящие названия колонок начинаются только с 3-й строки.
#    header=2 говорит pandas:
#    "используй 3-ю строку как заголовок таблицы".
#
# 4) Что такое None?
#    None - это "пустое значение" в Python (аналог пустой ячейки / NULL).
#    Если функция не может корректно преобразовать данные (например, телефон "abc"),
#    она возвращает None, чтобы показать:
#    "это значение невалидно и требует внимания".
#
# 5) Что делают apply(...) и функции norm_*?
#    Когда пишем:
#        df["phone"] = df["Телефон (исходный)"].apply(norm_phone)
#    это значит:
#    - берем каждую ячейку из колонки "Телефон (исходный)",
#    - пропускаем через функцию norm_phone,
#    - результат записываем в новую колонку "phone".
#
# 6) Почему мы не редактируем исходный файл напрямую?
#    Это безопаснее:
#    - исходник остается нетронутым,
#    - можно в любой момент перепроверить/переделать,
#    - результат сохраняется в отдельные "чистые" файлы.
#
# 7) Что значит "нормализация" данных?
#    Нормализация = привести разные форматы к единому стандарту.
#    Примеры:
#    - телефоны: "+7 (999) 111-22-33" -> "79991112233"
#    - даты: "15.01.2025" -> "2025-01-15"
#    - суммы: "1 250 ₽" -> 1250.0
#    - статусы: "Завершён" / "закрыт" -> "completed"
#
# 8) Почему в коде часто есть .copy()?
#    pandas иногда предупреждает, если менять "вид" таблицы напрямую.
#    .copy() создает независимую копию, с которой безопасно работать и изменять ее.
#
# 9) Как запускается скрипт?
#    В терминале, находясь в папке проекта:
#        python solve_lesson13.py
#    После запуска создаются:
#    - 5 отдельных .xlsx (по каждому заданию),
#    - 1 общий .xlsx с 5 листами.
#
# 10) Как понять, что все прошло успешно?
#     В конце работы скрипт печатает сообщение:
#     "Done: saved 5 separate files and 1 combined workbook ..."
#     Это означает, что обработка завершилась без критических ошибок.
#
# 11) Что делать, если появляется ошибка "ModuleNotFoundError: No module named pandas"?
#     Значит pandas не установлен в вашей среде Python.
#     Установите:
#         pip install pandas openpyxl
#     И потом снова запустите скрипт.
#
# 12) Зачем нужен openpyxl?
#     pandas использует openpyxl как "движок" для записи/чтения .xlsx.
#     Без него сохранение Excel-файлов может не работать.
#
# 13) Что делать, если данные обработались "не так, как ожидали"?
#     Откройте соответствующую функцию:
#     - телефон -> norm_phone
#     - дата -> norm_date
#     - сумма -> norm_amount
#     - статус -> norm_status
#     и посмотрите правила внутри (они подробно расписаны комментариями).
#
# 14) Почему часть строк может исчезнуть в итоге?
#     Это нормально для этапа очистки:
#     - служебные строки (ИТОГО, ---) удаляются,
#     - полностью невалидные строки удаляются,
#     - полные дубли помечаются как "дубль", ID у них пустой
#       (чтобы можно было удалить вручную после проверки).
#
# 15) Можно ли использовать этот файл как шаблон для других домашних?
#     Да. Обычно меняется:
#     - имя входного файла,
#     - названия листов и колонок,
#     - правила очистки в norm_* функциях.
#     Сама структура pipeline (read -> clean -> save) остается такой же.
