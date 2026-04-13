import os
import random
import time
import traceback
from collections import deque
from pathlib import Path

import requests
from requests import exceptions as req_exc
from dotenv import load_dotenv

from main_processor import process_excel_to_workbook


load_dotenv()

VK_GROUP_TOKEN = os.getenv("VK_GROUP_TOKEN", "")
VK_GROUP_ID = os.getenv("VK_GROUP_ID", "")
VK_API_VERSION = os.getenv("VK_API_VERSION", "5.199")
WORK_DIR = Path(os.getenv("WORK_DIR", "data"))


def vk_api(method: str, params: dict) -> dict:
    full_params = {
        **params,
        "access_token": VK_GROUP_TOKEN,
        "v": VK_API_VERSION,
    }
    response = requests.get(f"https://api.vk.com/method/{method}", params=full_params, timeout=30)
    response.raise_for_status()
    payload = response.json()
    if "error" in payload:
        raise RuntimeError(f"VK error in {method}: {payload['error']}")
    return payload["response"]


def send_text(peer_id: int, text: str) -> None:
    vk_api(
        "messages.send",
        {
            "peer_id": peer_id,
            "random_id": random.randint(1, 2_000_000_000),
            "message": text,
        },
    )


def send_file_to_user(peer_id: int, file_path: Path) -> None:
    upload_info = vk_api("docs.getMessagesUploadServer", {"peer_id": peer_id, "type": "doc"})
    with file_path.open("rb") as file_stream:
        upload_resp = requests.post(
            upload_info["upload_url"],
            files={"file": ("Ответ в ВК.xlsx", file_stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")},
            timeout=120,
        )
    upload_resp.raise_for_status()
    uploaded = upload_resp.json()["file"]

    saved = vk_api("docs.save", {"file": uploaded, "title": "Ответ в ВК.xlsx"})
    doc_obj = None
    if isinstance(saved, dict):
        doc_obj = saved.get("doc")
    elif isinstance(saved, list) and saved:
        first = saved[0]
        if isinstance(first, dict):
            doc_obj = first.get("doc", first)

    if not isinstance(doc_obj, dict) or "owner_id" not in doc_obj or "id" not in doc_obj:
        raise RuntimeError(f"Некорректный ответ docs.save: {saved}")

    attachment = f"doc{doc_obj['owner_id']}_{doc_obj['id']}"

    vk_api(
        "messages.send",
        {
            "peer_id": peer_id,
            "random_id": random.randint(1, 2_000_000_000),
            "message": "Готово: обработанный Excel-файл (5 листов).",
            "attachment": attachment,
        },
    )


def run() -> None:
    if not VK_GROUP_TOKEN or not VK_GROUP_ID:
        raise RuntimeError("VK_GROUP_TOKEN/VK_GROUP_ID не заданы в .env")

    WORK_DIR.mkdir(parents=True, exist_ok=True)
    try:
        long_poll = vk_api("groups.getLongPollServer", {"group_id": VK_GROUP_ID})
    except RuntimeError as exc:
        raise RuntimeError(
            "Не удалось запустить VK long poll. Проверьте, что токен группы выдан для этого сообщества "
            "и имеет права messages/manage."
        ) from exc
    server = long_poll["server"]
    key = long_poll["key"]
    ts = long_poll["ts"]
    recent_ids: deque[str] = deque()
    recent_ids_set: set[str] = set()
    print("VK worker запущен.")

    while True:
        try:
            poll_resp = requests.get(
                server,
                params={"act": "a_check", "key": key, "ts": ts, "wait": 25},
                timeout=35,
            )
            poll_resp.raise_for_status()
            payload = poll_resp.json()
        except (req_exc.Timeout, req_exc.ConnectionError) as exc:
            print(f"Сетевая ошибка VK long poll: {exc}")
            time.sleep(2)
            continue
        except req_exc.RequestException as exc:
            print(f"Ошибка VK long poll: {exc}")
            time.sleep(2)
            continue

        failed = payload.get("failed")
        if failed == 1:
            ts = payload.get("ts", ts)
            continue
        if failed in {2, 3}:
            long_poll = vk_api("groups.getLongPollServer", {"group_id": VK_GROUP_ID})
            server = long_poll["server"]
            key = long_poll["key"]
            ts = long_poll["ts"]
            continue

        ts = payload.get("ts", ts)

        for event in payload.get("updates", []):
            if event.get("type") != "message_new":
                continue
            message = event["object"]["message"]
            peer_id = message.get("peer_id")
            from_id = message.get("from_id")
            if from_id == -int(VK_GROUP_ID):
                continue

            message_id = message.get("id")
            conversation_message_id = message.get("conversation_message_id")
            dedupe_key = f"{peer_id}:{message_id}:{conversation_message_id}"
            if dedupe_key in recent_ids_set:
                continue
            if len(recent_ids) >= 200:
                oldest = recent_ids.popleft()
                recent_ids_set.discard(oldest)
            recent_ids.append(dedupe_key)
            recent_ids_set.add(dedupe_key)

            attachments = message.get("attachments", [])
            text = (message.get("text") or "").strip().lower()

            doc_info = None
            for item in attachments:
                if item.get("type") == "doc":
                    doc_info = item.get("doc")
                    break

            if not peer_id:
                continue

            if text == "/start" or (text and not doc_info):
                send_text(peer_id, "Пришлите Excel-файл (.xlsx или .xls) для обработки.")
                continue
            if not doc_info:
                send_text(peer_id, "Вам нужно прислать эксель файл.")
                continue

            url = doc_info.get("url")
            ext = str(doc_info.get("ext") or "").lower()
            if not url or ext not in {"xlsx", "xls"}:
                send_text(peer_id, "Вам нужно прислать эксель файл.")
                continue

            input_file = WORK_DIR / f"vk_{peer_id}_input.xlsx"
            output_file = WORK_DIR / f"vk_{peer_id}_processed.xlsx"

            try:
                source = requests.get(url, timeout=120)
                source.raise_for_status()
                input_file.write_bytes(source.content)

                process_excel_to_workbook(input_file, output_file)
                send_file_to_user(peer_id, output_file)
            except Exception as exc:
                print(f"Ошибка обработки VK файла: {exc!r}")
                traceback.print_exc()
                if "docs.getMessagesUploadServer" in str(exc):
                    send_text(
                        peer_id,
                        "Не хватает прав токена VK для отправки документов. "
                        "Включите доступ к документам сообщества в ключе доступа.",
                    )
                else:
                    send_text(peer_id, "Ошибка обработки файла. Проверьте, что вы отправили корректный Excel.")
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
