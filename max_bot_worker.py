import os
import time
from pathlib import Path

import requests
from requests import exceptions as req_exc
from dotenv import load_dotenv

from main_processor import process_excel_to_workbook


load_dotenv()

MAX_BOT_TOKEN = os.getenv("MAX_BOT_TOKEN", "")
WORK_DIR = Path(os.getenv("WORK_DIR", "data"))
MAX_API_BASE_URL = "https://platform-api.max.ru"


def headers() -> dict[str, str]:
    return {"Authorization": MAX_BOT_TOKEN}


def get_bot_user_id() -> int | None:
    response = requests.get(f"{MAX_API_BASE_URL}/me", headers=headers(), timeout=20)
    response.raise_for_status()
    payload = response.json()
    value = payload.get("user_id")
    return int(value) if isinstance(value, int) else None


def extract_chat_id(update: dict) -> int | None:
    message = update.get("message", {})
    recipient = message.get("recipient", {})
    return recipient.get("chat_id") or recipient.get("chatId")


def extract_text(update: dict) -> str:
    body = update.get("message", {}).get("body", {})
    return str(body.get("text") or "").strip().lower()


def find_file_attachment(update: dict) -> dict | None:
    attachments = update.get("message", {}).get("body", {}).get("attachments") or []
    for item in attachments:
        kind = str(item.get("type", "")).lower()
        payload = item.get("payload", {})
        file_name = str(payload.get("file_name") or payload.get("name") or "").lower()
        mime = str(payload.get("mime_type") or payload.get("mimeType") or "").lower()
        if kind in {"file", "document", "attachment"}:
            return item
        if file_name.endswith((".xlsx", ".xls")) or mime in {
            "application/vnd.ms-excel",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        }:
            return item
    return None


def attachment_download_url(attachment: dict) -> str | None:
    payload = attachment.get("payload", {})
    for key in ("url", "file_url", "download_url", "fileUrl", "downloadUrl"):
        value = payload.get(key)
        if value:
            return str(value)
    for key in ("url", "file_url", "download_url", "fileUrl", "downloadUrl"):
        value = attachment.get(key)
        if value:
            return str(value)
    return None


def send_text(chat_id: int, text: str) -> None:
    response = requests.post(
        f"{MAX_API_BASE_URL}/messages",
        headers={**headers(), "Content-Type": "application/json"},
        params={"chat_id": chat_id},
        json={"text": text},
        timeout=30,
    )
    response.raise_for_status()


def upload_file(file_path: Path) -> str:
    upload_init = requests.post(
        f"{MAX_API_BASE_URL}/uploads",
        headers=headers(),
        params={"type": "file"},
        timeout=30,
    )
    upload_init.raise_for_status()
    upload_data = upload_init.json()
    upload_url = upload_data.get("url")
    if not upload_url:
        raise RuntimeError(f"MAX uploads response has no url: {upload_data}")

    with file_path.open("rb") as stream:
        upload_resp = requests.post(
            upload_url,
            headers=headers(),
            files={"data": ("Ответ в MAX.xlsx", stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            timeout=120,
        )
    upload_resp.raise_for_status()
    upload_result = upload_resp.json()
    token = upload_result.get("token")
    if not token:
        raise RuntimeError(f"MAX upload response has no token: {upload_result}")
    return str(token)


def send_file_message(chat_id: int, file_token: str, text: str) -> None:
    body = {
        "text": text,
        "attachments": [
            {
                "type": "file",
                "payload": {"token": file_token, "file_name": "Ответ в MAX.xlsx"},
            }
        ],
    }
    response = requests.post(
        f"{MAX_API_BASE_URL}/messages",
        headers={**headers(), "Content-Type": "application/json"},
        params={"chat_id": chat_id},
        json=body,
        timeout=30,
    )
    response.raise_for_status()


def run() -> None:
    if not MAX_BOT_TOKEN:
        raise RuntimeError("MAX_BOT_TOKEN не задан в .env")

    WORK_DIR.mkdir(parents=True, exist_ok=True)
    bot_user_id = get_bot_user_id()
    marker = None
    print("MAX worker запущен.")

    while True:
        try:
            response = requests.get(
                f"{MAX_API_BASE_URL}/updates",
                headers=headers(),
                params={
                    "timeout": 30,
                    "types": "message_created",
                    "marker": marker,
                },
                timeout=40,
            )
            response.raise_for_status()
            payload = response.json()
        except (req_exc.Timeout, req_exc.ConnectionError) as exc:
            print(f"Сетевая ошибка MAX updates: {exc}")
            time.sleep(2)
            continue
        except req_exc.RequestException as exc:
            print(f"Ошибка MAX updates: {exc}")
            time.sleep(2)
            continue

        updates = payload.get("updates", []) or []
        marker = payload.get("marker", marker)

        for update in updates:
            sender = update.get("message", {}).get("sender", {}) or {}
            sender_user_id = sender.get("user_id")
            if sender.get("is_bot") is True:
                continue
            if bot_user_id is not None and sender_user_id == bot_user_id:
                continue

            chat_id = extract_chat_id(update)
            if not chat_id:
                continue

            text = extract_text(update)
            attachment = find_file_attachment(update)

            if text == "/start" or (text and not attachment):
                send_text(chat_id, "Пришлите Excel-файл (.xlsx или .xls) для обработки.")
                continue
            if not attachment:
                send_text(chat_id, "Вам нужно прислать эксель файл.")
                continue

            source_url = attachment_download_url(attachment)
            if not source_url:
                send_text(chat_id, "Не удалось получить ссылку на файл. Пришлите Excel заново.")
                continue

            input_file = WORK_DIR / f"max_{chat_id}_input.xlsx"
            output_file = WORK_DIR / f"max_{chat_id}_processed.xlsx"

            try:
                file_response = requests.get(source_url, headers=headers(), timeout=120)
                file_response.raise_for_status()
                input_file.write_bytes(file_response.content)

                process_excel_to_workbook(input_file, output_file)
                token = upload_file(output_file)
                time.sleep(1)
                send_file_message(chat_id, token, "Готово: обработанный Excel-файл (5 листов).")
            except Exception as exc:
                print(f"Ошибка обработки MAX файла: {exc}")
                send_text(chat_id, "Вам нужно прислать эксель файл.")
            finally:
                for temp_file in (input_file, output_file):
                    try:
                        if temp_file.exists():
                            temp_file.unlink()
                    except OSError as cleanup_error:
                        print(f"Не удалось удалить временный файл {temp_file}: {cleanup_error}")

        time.sleep(1)


if __name__ == "__main__":
    run()
