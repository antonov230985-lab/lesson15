import os
import time
from pathlib import Path

import requests
from dotenv import load_dotenv

from main_processor import process_csv


load_dotenv()

MAX_API_BASE_URL = os.getenv("MAX_API_BASE_URL", "").rstrip("/")
MAX_BOT_TOKEN = os.getenv("MAX_BOT_TOKEN", "")
MAX_CHAT_ID = os.getenv("MAX_CHAT_ID", "")
WORK_DIR = Path(os.getenv("WORK_DIR", "data"))


def headers() -> dict[str, str]:
    return {"Authorization": f"Bearer {MAX_BOT_TOKEN}"}


def run() -> None:
    if not MAX_API_BASE_URL or not MAX_BOT_TOKEN:
        raise RuntimeError("MAX_API_BASE_URL/MAX_BOT_TOKEN не заданы в .env")

    WORK_DIR.mkdir(parents=True, exist_ok=True)
    since_id = 0
    print("MAX worker запущен.")

    while True:
        response = requests.get(
            f"{MAX_API_BASE_URL}/bot/updates",
            headers=headers(),
            params={"since_id": since_id, "chat_id": MAX_CHAT_ID or None},
            timeout=30,
        )
        response.raise_for_status()
        payload = response.json()
        updates = payload.get("updates", [])

        for update in updates:
            since_id = max(since_id, int(update.get("id", 0)))
            attachment_url = update.get("file_url")
            chat_id = update.get("chat_id")
            if not attachment_url or not chat_id:
                continue

            input_file = WORK_DIR / f"max_{chat_id}_input.csv"
            output_file = WORK_DIR / f"max_{chat_id}_output.csv"

            file_response = requests.get(attachment_url, timeout=120)
            file_response.raise_for_status()
            input_file.write_bytes(file_response.content)

            process_csv(input_file, output_file)

            with output_file.open("rb") as file_stream:
                send_resp = requests.post(
                    f"{MAX_API_BASE_URL}/bot/send-file",
                    headers=headers(),
                    data={"chat_id": str(chat_id), "caption": "Файл обработан"},
                    files={"file": (output_file.name, file_stream, "text/csv")},
                    timeout=120,
                )
            send_resp.raise_for_status()

        time.sleep(2)


if __name__ == "__main__":
    run()
