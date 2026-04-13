import re
from pathlib import Path

import pandas as pd


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

STATUS_MAP = {
    "завершен": "completed",
    "завершён": "completed",
    "закрыт": "completed",
    "в работе": "in_progress",
    "отменен": "cancelled",
    "отменён": "cancelled",
}


def norm_phone(value: object) -> str | None:
    if pd.isna(value):
        return None

    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]

    digits = re.sub(r"\D", "", text)
    if len(digits) == 10:
        digits = "7" + digits
    elif len(digits) == 11 and digits.startswith("8"):
        digits = "7" + digits[1:]

    if len(digits) == 11 and digits.startswith("7"):
        return digits
    return None


def norm_date(value: object) -> str | None:
    if pd.isna(value):
        return None

    text = str(value).strip().lower()
    text = text.replace("-", " ").replace("/", " ").replace(".", " ")

    for ru, en in RU_MONTHS.items():
        text = re.sub(rf"\b{ru}\w*\b", en, text)

    text = re.sub(r"\s+", " ", text).strip()

    if re.match(r"^\d{4}\s+\d{1,2}\s+\d{1,2}$", text):
        dt = pd.to_datetime(text, errors="coerce", format="%Y %m %d")
        if pd.notna(dt):
            return dt.strftime("%Y-%m-%d")

    for dayfirst in (True, False):
        dt = pd.to_datetime(text, errors="coerce", dayfirst=dayfirst)
        if pd.notna(dt):
            return dt.strftime("%Y-%m-%d")
    return None


def norm_amount(value: object) -> float | None:
    if pd.isna(value):
        return None

    text = str(value).strip().lower()
    text = text.replace("₽", "").replace("р", "").replace(" ", "")
    if text in {"", "нетданных", "nan"}:
        return None

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        parts = text.split(",")
        if len(parts) == 2 and len(parts[1]) in {1, 2}:
            text = text.replace(",", ".")
        else:
            text = text.replace(",", "")

    try:
        return float(text)
    except ValueError:
        return None


def norm_status(value: object) -> str | None:
    if pd.isna(value):
        return None
    return STATUS_MAP.get(str(value).strip().lower())


def solve_task_1(source: Path) -> pd.DataFrame:
    df = pd.read_excel(source, sheet_name="1_Телефоны_и_даты", header=2)
    result = df[["ID", "Клиент"]].copy()
    result["phone"] = df["Телефон (исходный)"].apply(norm_phone)
    result["date"] = df["Дата (исходная)"].apply(norm_date)
    return result


def solve_task_2(source: Path) -> pd.DataFrame:
    df = pd.read_excel(source, sheet_name="2_Суммы_и_статусы", header=2)
    result = df[["ID", "Услуга"]].copy()
    result["amount"] = df["Сумма (исходная)"].apply(norm_amount)
    result["status"] = df["Статус (исходный)"].apply(norm_status)
    return result


def solve_task_3(source: Path) -> pd.DataFrame:
    df = pd.read_excel(source, sheet_name="3_Мусор_и_дубли", header=2)
    work = df.copy()
    work["id_num"] = pd.to_numeric(work["#"], errors="coerce")
    work["amount_num"] = work["Сумма"].apply(norm_amount)
    work["phone_norm"] = work["Телефон"].apply(norm_phone)
    work["date_norm"] = work["Дата визита"].apply(norm_date)

    clean = work[
        work["id_num"].notna()
        & work["date_norm"].notna()
        & work["Клиент"].notna()
        & work["Услуга"].notna()
    ].copy()

    exact_key = ["Клиент", "phone_norm", "date_norm", "Услуга", "amount_num"]
    clean["is_exact_duplicate"] = clean.duplicated(subset=exact_key, keep="first")

    uniq = clean[~clean["is_exact_duplicate"]].copy()
    uniq["visit_number_for_client"] = (
        uniq.sort_values("date_norm").groupby(["Клиент", "phone_norm"]).cumcount() + 1
    )
    clean = clean.merge(
        uniq[exact_key + ["visit_number_for_client"]],
        on=exact_key,
        how="left",
    )

    clean["duplicate_note"] = clean.apply(
        lambda row: (
            "дубль"
            if row["is_exact_duplicate"]
            else ("повторный визит" if row["visit_number_for_client"] > 1 else "первичный визит")
        ),
        axis=1,
    )
    clean["id_out"] = clean["#"]
    clean.loc[clean["is_exact_duplicate"], "id_out"] = None

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
    return result


def solve_task_4(source: Path) -> pd.DataFrame:
    df = pd.read_excel(source, sheet_name="4_Логика_заказов", header=2)
    df["order_id"] = pd.to_numeric(df["order_id"], errors="coerce")
    df["price_num"] = pd.to_numeric(df["Цена"], errors="coerce")
    df["qty_num"] = pd.to_numeric(df["Кол-во"], errors="coerce")

    rows: list[dict] = []
    for order_id, part in df.groupby("order_id", dropna=True):
        client_row = part[part["Тип строки"] == "Клиент"].head(1)
        date = norm_date(client_row["Дата"].iloc[0]) if not client_row.empty else None
        client = client_row["Клиент"].iloc[0] if not client_row.empty else None
        phone = norm_phone(client_row["Телефон"].iloc[0]) if not client_row.empty else None

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

    return pd.DataFrame(rows).sort_values("order_id")


def solve_task_5(source: Path) -> pd.DataFrame:
    df = pd.read_excel(source, sheet_name="5_Комплексное", header=2)
    work = df.copy()
    work["id_num"] = pd.to_numeric(work["#"], errors="coerce")
    work = work[work["id_num"].notna()].copy()

    work["date"] = work["Дата"].apply(norm_date)
    work["client"] = work["Клиент"].astype(str).str.strip().str.title()
    work["phone"] = work["Телефон"].apply(norm_phone)
    work["vin"] = work["VIN"].astype(str).str.strip()
    work["service"] = work["Услуга"].astype(str).str.strip()
    work["amount"] = work["Сумма"].apply(norm_amount)
    work["status"] = work["Статус"].apply(norm_status)

    work.loc[work["vin"].isin(["nan", "None"]), "vin"] = None
    work.loc[work["service"].isin(["nan", "None"]), "service"] = None

    clean = work[
        work["date"].notna()
        & work["client"].notna()
        & work["phone"].notna()
        & work["service"].notna()
    ].copy()
    clean = clean.drop_duplicates(
        subset=["date", "client", "phone", "vin", "service", "amount", "status"],
        keep="first",
    )
    return clean[["#", "date", "client", "phone", "vin", "service", "amount", "status"]]


def process_excel_to_workbook(source_excel: Path, out_workbook: Path) -> Path:
    task1_df = solve_task_1(source_excel)
    task2_df = solve_task_2(source_excel)
    task3_df = solve_task_3(source_excel)
    task4_df = solve_task_4(source_excel)
    task5_df = solve_task_5(source_excel)

    out_workbook.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_workbook, engine="openpyxl") as writer:
        task1_df.to_excel(writer, sheet_name="task1_phones_dates", index=False)
        task2_df.to_excel(writer, sheet_name="task2_amounts_statuses", index=False)
        task3_df.to_excel(writer, sheet_name="task3_garbage_duplicates", index=False)
        task4_df.to_excel(writer, sheet_name="task4_orders_aggregated", index=False)
        task5_df.to_excel(writer, sheet_name="task5_full_pipeline", index=False)
    return out_workbook
