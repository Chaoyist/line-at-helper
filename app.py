# app.py
# LINE Bot：
# 使用者輸入「7日內國內線統計表」→ 回覆指定的短網址，並附上摘要（由 Google Sheets gviz CSV 抓取）
# 使用者輸入「當日疏運統計表」→ 回覆指定的短網址

import os
import csv
import requests
import datetime
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage

# 時區（Python 3.9+ 內建 zoneinfo），用於顯示台灣時間日期
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
except Exception:
    ZoneInfo = None

app = Flask(__name__)

CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET")
CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")

line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN) if CHANNEL_ACCESS_TOKEN else None
handler = WebhookHandler(CHANNEL_SECRET) if CHANNEL_SECRET else None

# ---- 資料來源（固定分頁名稱 + 範圍，避免 gid 飄移） ----
FILE_ID = "1Nttc45OMeYl5SysfxWJ0B5qUu9Bo42Hx"
CSV_URL = (
    f"https://docs.google.com/spreadsheets/d/{FILE_ID}/gviz/tq?"
    "tqx=out:csv&sheet=%E7%B5%B1%E8%A8%881&range=CP2:CS32"
)

# 對應列（以你提供的 CSV 行號為準，1-based）
ROW_MAP = {
    "全航線": 31,
    "金門航線": 7,
    "澎湖航線": 13,
    "馬祖航線": 18,
    "本島航線": 23,
    "其他離島航線": 30,
}

# ---- 當日疏運統計表（國內線 D1:P38）----
CSV_DAILY_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1KTPwIgiqB2AOoQI4P_TySam0l12DO7wd"
    "/gviz/tq?tqx=out:csv&sheet=%E5%9C%8B%E5%85%A7%E7%B7%9A&range=D1:P38"
)

# A1 標記轉 0-based index，例如 'M19' -> (18, 12)
def _a1_to_index(a1: str) -> tuple[int, int]:
    a1 = a1.strip().upper()
    i = 0
    while i < len(a1) and a1[i].isalpha():
        i += 1
    col_letters = a1[:i]
    row_digits = a1[i:]
    if not col_letters or not row_digits.isdigit():
        raise ValueError(f"Invalid A1: {a1}")
    col_num = 0
    for ch in col_letters:
        col_num = col_num * 26 + (ord(ch) - ord('A') + 1)
    col_idx = col_num - 1
    row_idx = int(row_digits) - 1
    return (row_idx, col_idx)

def _get_a1(rows: list[list[str]], a1: str, default: str = "-") -> str:
    r, c = _a1_to_index(a1)
    if r < 0 or r >= len(rows):
        return default
    row = rows[r]
    if c < 0 or c >= len(row):
        return default
    return (row[c] or "").strip() or default

def fetch_daily_transport_summary() -> tuple[str, str, str]:
    """
    擷取「當日疏運統計表」摘要三值：
    本日表定架次=M19、已飛架次=M34、取消架次=M28。
    任何錯誤一律以 '-' 回傳避免中斷。
    """
    try:
        resp = requests.get(CSV_DAILY_URL, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
        resp.raise_for_status()
        txt = resp.text.strip()
        if txt.startswith("<!DOCTYPE html"):
            raise RuntimeError("CSV endpoint returned HTML (check sharing settings)")
        rows = list(csv.reader(txt.splitlines()))
        scheduled = _get_a1(rows, "M19", "-")
        flown = _get_a1(rows, "M34", "-")
        cancelled = _get_a1(rows, "M28", "-")
        return (scheduled, flown, cancelled)
    except Exception:
        return ("-", "-", "-")

def fetch_summary_text() -> str:
    """抓取 CSV 並依固定列組成摘要文字。若失敗，回傳提示字串。"""
    try:
        resp = requests.get(CSV_URL, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
        resp.raise_for_status()
        # 簡單防呆：若回 HTML，多半是權限/重導
        if resp.text.strip().startswith("<!DOCTYPE html"):
            raise RuntimeError("CSV endpoint returned HTML (check sharing or publish settings)")

        rows = list(csv.reader(resp.text.splitlines()))
        # CSV 欄序：CP, CQ, CR, CS → index 0..3
        def get_values(row_1_based: int):
            i = row_1_based - 1
            if i < 0 or i >= len(rows):
                return ("-", "-", "-", "-")
            r = rows[i]
            # 保護：欄位不足補 '-'
            vals = [(r[j].strip() if j < len(r) and r[j] is not None else "-") for j in range(4)]
            return tuple(vals)

        # 標題：昨日(YYYY/MM/DD)航班彙整摘要
        y = datetime.date.today() - datetime.timedelta(days=1)
        title = f"\n\n昨日({y.strftime('%Y/%m/%d')})航班彙整摘要"

        parts = []
        parts.append(title)
        parts.append("全航線：")
        cp, cq, cr, cs = get_values(ROW_MAP["全航線"])
        parts.append(f"✈️ 架次：{cp}")
        parts.append(f"💺 座位數：{cq}")
        parts.append(f"👥 載客數：{cr}")
        parts.append(f"📊 載客率：{cs}")

        for route in ["金門航線", "澎湖航線", "馬祖航線", "本島航線", "其他離島航線"]:
            cp, cq, cr, cs = get_values(ROW_MAP[route])
            parts.append(f"\n{route}：")
            parts.append(f"✈️ 架次：{cp}")
            parts.append(f"💺 座位數：{cq}")
            parts.append(f"👥 載客數：{cr}")
            parts.append(f"📊 載客率：{cs}")

        return "\n".join(parts)
    except Exception as e:
        return f"（暫時無法取得統計資料：{e}）"

# ---- LINE Webhook ----
@app.route("/callback", methods=["POST"])
def callback():
    if not handler:
        return "LINE handler not configured", 503

    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"

# ---- 訊息處理 ----
if handler:
    @handler.add(MessageEvent, message=TextMessage)
    def handle_message(event: MessageEvent):
        text = (event.message.text or "").strip()

        if text == "7日內國內線統計表":
            url = "https://reurl.cc/Lnrjdy"
            summary = fetch_summary_text()
            msg = f"📈 7日內國內線統計表：\n{url}{summary and ('' + summary)}"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
            return

        if text == "當日疏運統計表":
            url = "https://reurl.cc/9nNEAO"
            scheduled, flown, cancelled = fetch_daily_transport_summary()
            # 以台灣時區顯示今天日期
            if ZoneInfo:
                today = datetime.datetime.now(ZoneInfo("Asia/Taipei")).strftime("%Y/%m/%d")
            else:
                today = datetime.datetime.now().strftime("%Y/%m/%d")
            msg = (
                f"📊 當日疏運統計表：{url}"
                f"\n摘要 ({today})"
                f"\n本日表定架次：{scheduled}"
                f"\n已飛架次：{flown}"
                f"\n取消架次：{cancelled}"
            )
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
            return

        tip = "請輸入「7日內國內線統計表」或「當日疏運統計表」🙂"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=tip))

# ---- 根路由 ----
@app.route("/", methods=["GET"])
def index():
    return ("Flight Bot online. POST to /callback with LINE events", 200)
