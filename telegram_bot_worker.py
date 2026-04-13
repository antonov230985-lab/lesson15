import os
import time
from pathlib import Path

import requests
from dotenv import load_dotenv

from main_processor import process_csv


load_dotenv()

TELEGRAM_BOT_TOKEN = os.getenv("TELEGRAM_BOT_TOKEN", "")
TELEGRAM_CHAT_ID = os.getenv("TELEGRAM_CHAT_ID", "")
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


def send_document(chat_id: int | str, csv_path: Path) -> None:
    with csv_path.open("rb") as file_stream:
        response = requests.post(
            tg_api_url("sendDocument"),
            data={"chat_id": str(chat_id), "caption": "Файл обработан"},
            files={"document": (csv_path.name, file_stream, "text/csv")},
            timeout=120,
        )
    response.raise_for_status()


def run() -> None:
    if not TELEGRAM_BOT_TOKEN:
        raise RuntimeError("TELEGRAM_BOT_TOKEN не задан в .env")

    allowed_chat = str(TELEGRAM_CHAT_ID).strip() if TELEGRAM_CHAT_ID else None
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
            if not document or chat_id is None:
                continue

            if allowed_chat and str(chat_id) != allowed_chat:
                continue

            input_file = WORK_DIR / f"telegram_{chat_id}_input.csv"
            output_file = WORK_DIR / f"telegram_{chat_id}_output.csv"

            download_document(document["file_id"], input_file)
            process_csv(input_file, output_file)
            send_document(chat_id, output_file)


if __name__ == "__main__":
    run()
