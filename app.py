# app.py
# LINE Bot：
# 使用者輸入「7日內國內線統計表」→ 回覆指定的短網址，並附上摘要（由 Google Sheets gviz CSV 抓取）
# 使用者輸入「當日疏運統計表」→ 回覆指定的短網址

import os
import csv
import requests
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage

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
    "全航線": 30,
    "金門航線": 6,
    "澎湖航線": 12,
    "馬祖航線": 17,
    "本島航線": 22,
    "其他離島航線": 29,
}

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

        parts = []
        parts.append("全航線：")
        cp, cq, cr, cs = get_values(ROW_MAP["全航線"])
        parts.append(f"✈️ 架次：{cp}")
        parts.append(f"💺 座位數：{cq}")
        parts.append(f"👥 載客數：{cr}")
        parts.append(f"📊 載客率：{cs}")

        for route in ["金門航線", "澎湖航線", "馬祖航線", "本島航線", "其他離島航線"]:
            cp, cq, cr, cs = get_values(ROW_MAP[route])
            parts.append(f"{route}：")
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
            msg = f"📈 7日內國內線統計表：{url}{summary and ('' + summary)}"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
            return

        if text == "當日疏運統計表":
            url = "https://reurl.cc/9nNEAO"
            msg = f"📊 當日疏運統計表：\n{url}"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
            return

        tip = "請輸入「7日內國內線統計表」或「當日疏運統計表」🙂"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=tip))

# ---- 根路由 ----
@app.route("/", methods=["GET"])
def index():
    return ("Flight Bot online. POST to /callback with LINE events", 200)
