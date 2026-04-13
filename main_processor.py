import argparse
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


def process_csv(input_path: Path, output_path: Path) -> Path:
    df = pd.read_csv(input_path)
    cleaned = df.copy()

    for column in cleaned.columns:
        normalized = column.strip().lower()
        if "телефон" in normalized or normalized in {"phone", "phone_raw"}:
            cleaned[column] = cleaned[column].apply(norm_phone)
        elif "дата" in normalized or normalized in {"date", "visit_date"}:
            cleaned[column] = cleaned[column].apply(norm_date)
        elif "сумм" in normalized or normalized in {"amount", "amount_raw", "price"}:
            cleaned[column] = cleaned[column].apply(norm_amount)
        elif "статус" in normalized or normalized == "status":
            cleaned[column] = cleaned[column].apply(norm_status)

    cleaned = cleaned.drop_duplicates().reset_index(drop=True)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    cleaned.to_csv(output_path, index=False, encoding="utf-8")
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Очистка входного CSV-файла.")
    parser.add_argument("--input", required=True, help="Путь к входному CSV")
    parser.add_argument("--output", required=True, help="Путь к выходному CSV")
    args = parser.parse_args()

    output_path = process_csv(Path(args.input), Path(args.output))
    print(f"Готово: {output_path}")


if __name__ == "__main__":
    main()
