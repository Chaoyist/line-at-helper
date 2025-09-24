# app.py
# LINE Bot：使用者輸入「七日內國內線統計報表」→ 回覆指定的短網址，並附上昨日的摘要統計（從 Google Sheets 抓取）

import os
import requests
import datetime
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage

app = Flask(__name__)

CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET")
CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")

line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN) if CHANNEL_ACCESS_TOKEN else None
handler = WebhookHandler(CHANNEL_SECRET) if CHANNEL_SECRET else None

# Google Sheets CSV 匯出連結 (僅讀，不會動到原始檔)
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1Nttc45OMeYl5SysfxWJ0B5qUu9Bo42Hx/export?format=csv&gid=1842879320"

# ---- 讀取 Google Sheets 並整理昨日資料 ----
def fetch_yesterday_summary():
    try:
        # 計算昨日日期 (假設今天 2025/09/24 → 昨日 2025/09/23)
        today = datetime.date.today()
        yesterday = today - datetime.timedelta(days=1)
        target_date = yesterday.strftime("%Y/%m/%d")

        # 下載 CSV 內容
        resp = requests.get(SHEET_CSV_URL)
        resp.raise_for_status()
        lines = resp.text.splitlines()

        # 假設表格格式：日期, 航線, 架次, 座位數, 載客數, 載客率
        import csv
        reader = csv.DictReader(lines)

        summary_texts = []
        for row in reader:
            if row.get("日期") == target_date and "金門" in row.get("航線", ""):
                summary_texts.append(
                    f"\n金門航線：\n"
                    f"✈️ 架次：{row.get('架次','-')}\n"
                    f"💺 座位數：{row.get('座位數','-')}\n"
                    f"👥 載客數：{row.get('載客數','-')}\n"
                    f"📊 載客率：{row.get('載客率','-')}"
                )

        if summary_texts:
            return f"昨日({target_date})國內線簡要統計" + "".join(summary_texts)
        else:
            return f"昨日({target_date})國內線簡要統計：查無資料"

    except Exception as e:
        return f"無法讀取昨日統計資料：{e}"

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

        if text == "七日內國內線統計報表":
            url = "https://reurl.cc/Lnrjdy"
            summary = fetch_yesterday_summary()
            msg = f"📈 七日內國內線統計報表：\n{url}\n\n{summary}"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
            return

        tip = "請輸入「七日內國內線統計報表」🙂"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=tip))

# ---- 根路由 ----
@app.route("/", methods=["GET"])
def index():
    return ("Flight Bot online. POST to /callback with LINE events", 200)
