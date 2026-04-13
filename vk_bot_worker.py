import os
import random
import time
from pathlib import Path

import requests
from dotenv import load_dotenv

from main_processor import process_csv


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


def send_file_to_user(peer_id: int, csv_path: Path) -> None:
    upload_info = vk_api("docs.getMessagesUploadServer", {"peer_id": peer_id, "type": "doc"})
    with csv_path.open("rb") as file_stream:
        upload_resp = requests.post(
            upload_info["upload_url"],
            files={"file": (csv_path.name, file_stream, "text/csv")},
            timeout=120,
        )
    upload_resp.raise_for_status()
    uploaded = upload_resp.json()["file"]

    saved = vk_api("docs.save", {"file": uploaded, "title": csv_path.name})
    attachment = f"doc{saved['doc']['owner_id']}_{saved['doc']['id']}"

    vk_api(
        "messages.send",
        {
            "peer_id": peer_id,
            "random_id": random.randint(1, 2_000_000_000),
            "message": "Файл обработан",
            "attachment": attachment,
        },
    )


def run() -> None:
    if not VK_GROUP_TOKEN or not VK_GROUP_ID:
        raise RuntimeError("VK_GROUP_TOKEN/VK_GROUP_ID не заданы в .env")

    WORK_DIR.mkdir(parents=True, exist_ok=True)
    long_poll = vk_api("groups.getLongPollServer", {"group_id": VK_GROUP_ID})
    server = long_poll["server"]
    key = long_poll["key"]
    ts = long_poll["ts"]
    print("VK worker запущен.")

    while True:
        poll_resp = requests.get(
            server,
            params={"act": "a_check", "key": key, "ts": ts, "wait": 25},
            timeout=35,
        )
        poll_resp.raise_for_status()
        payload = poll_resp.json()
        ts = payload.get("ts", ts)

        for event in payload.get("updates", []):
            if event.get("type") != "message_new":
                continue
            message = event["object"]["message"]
            peer_id = message.get("peer_id")
            attachments = message.get("attachments", [])

            doc_info = None
            for item in attachments:
                if item.get("type") == "doc":
                    doc_info = item.get("doc")
                    break
            if not doc_info or not peer_id:
                continue

            url = doc_info.get("url")
            if not url or not doc_info.get("ext", "").lower() == "csv":
                continue

            input_file = WORK_DIR / f"vk_{peer_id}_input.csv"
            output_file = WORK_DIR / f"vk_{peer_id}_output.csv"

            source = requests.get(url, timeout=120)
            source.raise_for_status()
            input_file.write_bytes(source.content)

            process_csv(input_file, output_file)
            send_file_to_user(peer_id, output_file)

        time.sleep(1)


if __name__ == "__main__":
    run()
