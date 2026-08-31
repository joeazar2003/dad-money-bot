import os
import re
import io
from collections import defaultdict
from datetime import datetime

import requests
from flask import Flask, request
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font

MONTH_NAMES = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
]

app = Flask(__name__)

TELEGRAM_TOKEN = os.environ["TELEGRAM_TOKEN"]
# Your personal Telegram numeric chat ID. Only messages from this chat are logged.
ALLOWED_CHAT_ID = int(os.environ["ALLOWED_CHAT_ID"])

API_URL = f"https://api.telegram.org/bot{TELEGRAM_TOKEN}"
XLSX_PATH = "dad_money.xlsx"

HELP_TEXT = (
    "Hey! Just text me an amount whenever your dad sends you money, e.g.:\n\n"
    "  5000 from Bassam Hanna\n"
    "  five thousand dollars from Serge\n\n"
    "Commands:\n"
    "/total - see your running total\n"
    "/download - get the Excel file\n"
    "/undo - remove the last entry\n\n"
    "Edited the Excel file yourself? Just send it back to me as a file "
    "attachment and I'll use your edited version going forward.\n"
)

AMOUNT_RE = re.compile(r"[-+]?\d+(?:\.\d+)?")

NUMBER_WORDS = {
    "zero": 0, "one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
    "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11, "twelve": 12,
    "thirteen": 13, "fourteen": 14, "fifteen": 15, "sixteen": 16, "seventeen": 17,
    "eighteen": 18, "nineteen": 19, "twenty": 20, "thirty": 30, "forty": 40,
    "fifty": 50, "sixty": 60, "seventy": 70, "eighty": 80, "ninety": 90,
}
SCALE_WORDS = {"hundred": 100, "thousand": 1000, "million": 1000000, "billion": 1000000000}
FILLER_AFTER_AMOUNT = {"dollars", "dollar", "usd", "bucks"}


def words_to_number(words):
    total = 0
    current = 0
    for w in words:
        if w in NUMBER_WORDS:
            current += NUMBER_WORDS[w]
        elif w in SCALE_WORDS:
            scale = SCALE_WORDS[w]
            current = (current or 1) * scale
            if scale >= 1000:
                total += current
                current = 0
    return total + current


def clean_word(w):
    return re.sub(r"[^a-zA-Z]", "", w).lower()


def extract_amount_and_note(text):
    # 1) Prefer plain digits, e.g. "5000 from Bassam Hanna"
    match = AMOUNT_RE.search(text)
    if match:
        amount = float(match.group())
        note = (text[:match.start()] + text[match.end():]).strip(" $-")
        return amount, note

    # 2) Fall back to spelled-out numbers, e.g. "five thousand dollars from Bassam Hanna"
    words = text.split()
    n = len(words)
    best_start = best_end = None
    i = 0
    while i < n:
        if clean_word(words[i]) in NUMBER_WORDS or clean_word(words[i]) in SCALE_WORDS:
            start = i
            j = i
            while j < n and (clean_word(words[j]) in NUMBER_WORDS or clean_word(words[j]) in SCALE_WORDS):
                j += 1
            if best_start is None or (j - start) > (best_end - best_start):
                best_start, best_end = start, j
            i = j
        else:
            i += 1

    if best_start is None:
        return None, text

    number_words = [clean_word(w) for w in words[best_start:best_end]]
    amount = words_to_number(number_words)
    if amount <= 0:
        return None, text

    remaining = words[:best_start] + words[best_end:]
    if remaining and clean_word(remaining[0]) in FILLER_AFTER_AMOUNT:
        remaining = remaining[1:]
    note = " ".join(remaining).strip(" $-")
    return amount, note


def ensure_workbook():
    if not os.path.exists(XLSX_PATH):
        wb = Workbook()
        ws = wb.active
        ws.title = "Log"
        ws.append(["Date", "Amount", "Note"])
        wb.save(XLSX_PATH)


def rebuild_summary(wb):
    log_ws = wb["Log"]
    monthly = defaultdict(lambda: defaultdict(float))

    for row in log_ws.iter_rows(min_row=2, values_only=True):
        date_str, amount = row[0], row[1]
        if not date_str or amount is None:
            continue
        try:
            dt = datetime.strptime(str(date_str), "%Y-%m-%d %H:%M")
        except ValueError:
            continue
        monthly[dt.year][dt.month] += float(amount)

    if "Summary" in wb.sheetnames:
        del wb["Summary"]
    ws = wb.create_sheet("Summary")
    ws.append(["Year", "Month", "Total"])
    ws["A1"].font = ws["B1"].font = ws["C1"].font = Font(bold=True)

    for year in sorted(monthly.keys()):
        year_total = 0.0
        for month in sorted(monthly[year].keys()):
            total = monthly[year][month]
            year_total += total
            ws.append([year, MONTH_NAMES[month - 1], round(total, 2)])
        total_row = ws.max_row + 1
        ws.append([year, "Year Total", round(year_total, 2)])
        for col in ("A", "B", "C"):
            ws[f"{col}{total_row}"].font = Font(bold=True)
        ws.append([])

    for col, width in zip("ABC", (8, 14, 14)):
        ws.column_dimensions[col].width = width


def append_entry(amount, note):
    ensure_workbook()
    wb = load_workbook(XLSX_PATH)
    ws = wb["Log"]
    ws.append([datetime.now().strftime("%Y-%m-%d %H:%M"), amount, note])
    rebuild_summary(wb)
    wb.save(XLSX_PATH)


def undo_last():
    ensure_workbook()
    wb = load_workbook(XLSX_PATH)
    ws = wb["Log"]
    if ws.max_row <= 1:
        return None
    last_row = [c.value for c in ws[ws.max_row]]
    ws.delete_rows(ws.max_row)
    rebuild_summary(wb)
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


def handle_incoming_document(chat_id, document):
    file_name = document.get("file_name", "")
    if not file_name.lower().endswith(".xlsx"):
        send_message(chat_id, "That doesn't look like an .xlsx file, so I left your data untouched.")
        return

    file_id = document.get("file_id")
    info = requests.get(f"{API_URL}/getFile", params={"file_id": file_id}).json()
    if not info.get("ok"):
        send_message(chat_id, "Couldn't fetch that file from Telegram, try sending it again.")
        return

    file_path = info["result"]["file_path"]
    file_url = f"https://api.telegram.org/file/bot{TELEGRAM_TOKEN}/{file_path}"
    resp = requests.get(file_url)

    tmp_path = XLSX_PATH + ".tmp"
    with open(tmp_path, "wb") as f:
        f.write(resp.content)

    try:
        wb = load_workbook(tmp_path)
        if "Log" not in wb.sheetnames:
            raise ValueError("Missing Log sheet")
    except Exception as e:
        os.remove(tmp_path)
        send_message(chat_id, f"Couldn't read that file: {type(e).__name__}: {e}")
        return

    os.replace(tmp_path, XLSX_PATH)
    total, count = get_total()
    send_message(chat_id, f"Got it — saved your edits. Running total is now ${total:,.2f} ({count} entries). New entries will build on top of this.")


@app.route("/webhook", methods=["POST"])
def webhook():
    update = request.get_json(force=True, silent=True) or {}
    message = update.get("message", {})
    chat_id = message.get("chat", {}).get("id")

    if chat_id is None:
        return "ok"

    if chat_id != ALLOWED_CHAT_ID:
        # Ignore anyone who isn't you.
        return "ok"

    if "document" in message:
        handle_incoming_document(chat_id, message["document"])
        return "ok"

    text = (message.get("text") or "").strip()
    if not text:
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

    amount, note = extract_amount_and_note(text)
    if amount is None:
        send_message(chat_id, "Didn't catch an amount there. Try something like: 5000 from Bassam Hanna or five thousand dollars from Serge")
        return "ok"

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
