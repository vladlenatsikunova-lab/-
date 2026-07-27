"""
Ежедневный дайджест контент-плана в Telegram.

Берёт данные из Google Sheets (три листа: ЯМАЛМОТО, Геохакинг, Дома скучно),
находит строки с сегодняшней датой (колонка A) и присылает сообщение в Telegram
со всем, что запланировано на сегодня.

Нужны переменные окружения:
  TELEGRAM_BOT_TOKEN — токен бота от @BotFather
  TELEGRAM_CHAT_ID   — id чата/канала или @username публичного канала

Запуск: python telegram_daily.py
"""

import csv
import io
import json
import os
import re
import urllib.parse
import urllib.request
from datetime import datetime, timedelta, timezone

SHEET_ID = "1scpbz_mMTfhSC7q9NDgoNsIi19mRn-GInoDrxm7CLs4"
BOT_TOKEN = os.environ["TELEGRAM_BOT_TOKEN"]
CHAT_ID = os.environ["TELEGRAM_CHAT_ID"]

MONTHS_RU = {
    "января": 1, "февраля": 2, "марта": 3, "апреля": 4,
    "мая": 5, "июня": 6, "июля": 7, "августа": 8,
    "сентября": 9, "октября": 10, "ноября": 11, "декабря": 12,
}

SHEETS = ["ЯМАЛМОТО", "Геохакинг", "Дома скучно"]
EMOJI = {"ЯМАЛМОТО": "🏍", "Геохакинг": "🧭", "Дома скучно": "🏠"}

# значения, которые не считаются реальным контентом (чекбоксы, прочерки и т.п.)
IGNORE_VALUES = {"", "true", "false", "-", "—"}


def fetch_sheet_csv(sheet_name: str):
    url = (
        f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/gviz/tq"
        f"?tqx=out:csv&sheet={urllib.parse.quote(sheet_name)}"
    )
    with urllib.request.urlopen(url, timeout=30) as resp:
        raw = resp.read().decode("utf-8")
    return list(csv.reader(io.StringIO(raw)))


def parse_ru_date(cell: str):
    """'27 июля' -> (27, 7) или None"""
    m = re.match(r"^\s*(\d{1,2})\s+([а-яёА-ЯЁ]+)", cell.strip())
    if not m:
        return None
    day = int(m.group(1))
    month = MONTHS_RU.get(m.group(2).lower())
    if month is None:
        return None
    return day, month


def parse_full_date(cell: str):
    """'27.07.2026' -> date или None"""
    try:
        return datetime.strptime(cell.strip(), "%d.%m.%Y").date()
    except ValueError:
        return None


def clean(value: str):
    v = value.strip()
    if v.lower() in IGNORE_VALUES:
        return None
    return v


def today_msk():
    return (datetime.now(timezone.utc) + timedelta(hours=3)).date()


def collect_today_content(rows, sheet_name: str, today):
    items = []
    for row in rows:
        if not row:
            continue
        date_cell = row[0]
        if sheet_name == "Дома скучно":
            d = parse_full_date(date_cell)
            matched = d == today
        else:
            parsed = parse_ru_date(date_cell)
            matched = parsed == (today.day, today.month)
        if not matched:
            continue
        for cell in row[1:]:
            v = clean(cell)
            if v and v not in items:
                items.append(v)
    return items


def build_message(today):
    lines = [f"\U0001F4C5 {today.strftime('%d.%m.%Y')} — контент-план на сегодня\n"]
    any_content = False
    for sheet_name in SHEETS:
        try:
            rows = fetch_sheet_csv(sheet_name)
        except Exception as e:
            lines.append(f"{EMOJI[sheet_name]} {sheet_name}: ошибка загрузки ({e})\n")
            continue
        items = collect_today_content(rows, sheet_name, today)
        lines.append(f"{EMOJI[sheet_name]} {sheet_name}:")
        if items:
            any_content = True
            for it in items:
                lines.append(f"- {it}")
        else:
            lines.append("нет публикаций сегодня")
        lines.append("")
    if not any_content:
        lines.append("На сегодня по всем аккаунтам публикаций не найдено.")
    return "\n".join(lines).strip()


def send_telegram(text: str):
    url = f"https://api.telegram.org/bot{BOT_TOKEN}/sendMessage"
    data = urllib.parse.urlencode({
        "chat_id": CHAT_ID,
        "text": text,
        "disable_web_page_preview": "true",
    }).encode()
    req = urllib.request.Request(url, data=data, method="POST")
    with urllib.request.urlopen(req, timeout=30) as resp:
        result = json.loads(resp.read().decode())
    if not result.get("ok"):
        raise RuntimeError(f"Telegram API error: {result}")
    return result


if __name__ == "__main__":
    today = today_msk()
    message = build_message(today)
    print(message)
    send_telegram(message)
