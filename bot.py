import os
import re
import io
from datetime import datetime

import requests
from flask import Flask, request
from openpyxl import Workbook, load_workbook

app = Flask(__name__)

TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
# Your personal Telegram numeric chat ID. Only messages from this chat are logged.
ALLOWED_CHAT_ID = int(os.environ["ALLOWED_CHAT_ID"])

API_URL = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}"
XLSX_PATH = "dad_money.xlsx"

HELP_TEXT = (
    "Hey! Just text me an amount whenever your dad sends you money, e.g.:\n\n"
    "  50\n"
    "  120.50 birthday gift\n"
    "  $75 rent help\n\n"
    "Commands:\n"
    "/total - see your running total\n"
    "/download - get the Excel file\n"
    "/undo - remove the last entry\n"
)

AMOUNT_RE = re.compile(r"[-+]?\d+(?:\.\d+)?")


def ensure_workbook():
    if not os.path.exists(XLSX_PATH):
        wb = Workbook()
        ws = wb.active
        ws.title = "Log"
        ws.append(["Date", "Amount", "Note"])
        wb.save(XLSX_PATH)


def append_entry(amount, note):
    ensure_workbook()
    wb = load_workbook(XLSX_PATH)
    ws = wb["Log"]
    ws.append([datetime.now().strftime("%Y-%m-%d %H:%M"), amount, note])
    wb.save(XLSX_PATH)


def undo_last():
    ensure_workbook()
    wb = load_workbook(XLSX_PATH)
    ws = wb["Log"]
    if ws.max_row <= 1:
        return None
    last_row = [c.value for c in ws[ws.max_row]]
    ws.delete_rows(ws.max_row)
    wb.save(XLSX_PATH)
    return last_row


def get_total():
    ensure_workbook()
    wb = load_workbook(XLSX_PATH)
    ws = wb["Log"]
    total = 0.0
    count = 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[1] is not None:
            total += float(row[1])
            count += 1
    return total, count


def send_message(chat_id, text):
    requests.post(f"{API_URL}/sendMessage", json={"chat_id": chat_id, "text": text})


def send_document(chat_id):
    ensure_workbook()
    with open(XLSX_PATH, "rb") as f:
        requests.post(
            f"{API_URL}/sendDocument",
            data={"chat_id": chat_id},
            files={"document": ("dad_money.xlsx", f)},
        )


@app.route("/webhook", methods=["POST"])
def webhook():
    update = request.get_json(force=True, silent=True) or {}
    message = update.get("message", {})
    chat_id = message.get("chat", {}).get("id")
    text = (message.get("text") or "").strip()

    if chat_id is None or not text:
        return "ok"

    if chat_id != ALLOWED_CHAT_ID:
        # Ignore anyone who isn't you.
        return "ok"

    if text in ("/start", "/help"):
        send_message(chat_id, HELP_TEXT)
        return "ok"

    if text == "/total":
        total, count = get_total()
        send_message(chat_id, f"Total from dad: ${total:,.2f} across {count} entries.")
        return "ok"

    if text == "/download":
        send_document(chat_id)
        return "ok"

    if text == "/undo":
        removed = undo_last()
        if removed:
            send_message(chat_id, f"Removed: {removed[0]} — ${removed[1]} ({removed[2] or 'no note'})")
        else:
            send_message(chat_id, "Nothing to undo.")
        return "ok"

    match = AMOUNT_RE.search(text)
    if not match:
        send_message(chat_id, "Didn't catch an amount there. Try something like: 50 or $75 groceries")
        return "ok"

    amount = float(match.group())
    note = (text[:match.start()] + text[match.end():]).strip(" $-")
    append_entry(amount, note)
    total, count = get_total()
    send_message(chat_id, f"Logged ${amount:,.2f}. Running total: ${total:,.2f} ({count} entries).")
    return "ok"


@app.route("/", methods=["GET"])
def health():
    return "Bot is running."


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
