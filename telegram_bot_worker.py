import os
import time
from pathlib import Path

import requests
from dotenv import load_dotenv

from main_processor import process_excel_to_workbook


load_dotenv()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
WORK_DIR = Path(os.getenv("WORK_DIR", "data"))


def tg_api_url(method: str) -> str:
    return f"https://api.telegram.org/bot{TELEGRAM_BOT_TOKEN}/{method}"


def download_document(file_id: str, target_path: Path) -> None:
    file_info = requests.get(tg_api_url("getFile"), params={"file_id": file_id}, timeout=30).json()
    if not file_info.get("ok"):
        raise RuntimeError(f"Telegram getFile error: {file_info}")
    file_path = file_info["result"]["file_path"]

    file_url = f"https://api.telegram.org/file/bot{TELEGRAM_BOT_TOKEN}/{file_path}"
    response = requests.get(file_url, timeout=120)
    response.raise_for_status()
    target_path.write_bytes(response.content)


def send_message(chat_id: int | str, text: str) -> None:
    response = requests.post(
        tg_api_url("sendMessage"),
        data={"chat_id": str(chat_id), "text": text},
        timeout=30,
    )
    response.raise_for_status()


def send_document(chat_id: int | str, file_path: Path, caption: str = "Файл обработан") -> None:
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    reply_name = "Ответ в Телеграмм.xlsx"
    with file_path.open("rb") as file_stream:
        response = requests.post(
            tg_api_url("sendDocument"),
            data={"chat_id": str(chat_id), "caption": caption},
            files={"document": (reply_name, file_stream, mime)},
            timeout=120,
        )
    response.raise_for_status()


def is_excel_document(document: dict) -> bool:
    file_name = str(document.get("file_name", "")).lower()
    mime = str(document.get("mime_type", "")).lower()
    return file_name.endswith((".xlsx", ".xls")) or mime in {
        "application/vnd.ms-excel",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    }


def run() -> None:
    if not TELEGRAM_BOT_TOKEN:
        raise RuntimeError("TELEGRAM_BOT_TOKEN не задан в .env")

    WORK_DIR.mkdir(parents=True, exist_ok=True)
    offset = 0
    print("Telegram worker запущен.")

    while True:
        resp = requests.get(
            tg_api_url("getUpdates"),
            params={"timeout": 30, "offset": offset},
            timeout=40,
        )
        resp.raise_for_status()
        data = resp.json()
        if not data.get("ok"):
            time.sleep(2)
            continue

        for update in data.get("result", []):
            offset = update["update_id"] + 1
            message = update.get("message", {})
            chat_id = message.get("chat", {}).get("id")
            document = message.get("document")
            text = (message.get("text") or "").strip().lower()

            if chat_id is None:
                continue

            if text == "/start" or (text and not document):
                send_message(chat_id, "Пришлите Excel-файл (.xlsx или .xls) для обработки.")
                continue

            if not document:
                send_message(chat_id, "Вам нужно прислать эксель файл.")
                continue

            if not is_excel_document(document):
                send_message(chat_id, "Вам нужно прислать эксель файл.")
                continue

            input_file = WORK_DIR / f"telegram_{chat_id}_input.xlsx"
            output_file = WORK_DIR / f"telegram_{chat_id}_processed.xlsx"

            try:
                download_document(document["file_id"], input_file)
                process_excel_to_workbook(input_file, output_file)
                send_document(chat_id, output_file, caption="Готово: обработанный Excel-файл (5 листов).")
            except Exception as exc:
                print(f"Ошибка обработки Telegram файла: {exc}")
                send_message(chat_id, "Ошибка обработки файла. Проверьте, что вы отправили корректный Excel.")
            finally:
                for temp_file in (input_file, output_file):
                    try:
                        if temp_file.exists():
                            temp_file.unlink()
                    except OSError as cleanup_error:
                        print(f"Не удалось удалить временный файл {temp_file}: {cleanup_error}")


if __name__ == "__main__":
    run()
