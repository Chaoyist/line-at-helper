# app.py
# 簡易 LINE Bot：收到「選單」→ 回傳 Flex 圖文選單（兩個按鈕）
# 使用者按鈕觸發 postback → 回覆假的 Excel 下載網址

import os
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage,
    PostbackEvent, FlexSendMessage
)

app = Flask(__name__)

CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET")
CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")

line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(CHANNEL_SECRET)

# --- Flex 圖文選單（兩個選項） ---
def build_menu_flex():
    # 你可以換 hero 的示意圖片 URL
    flex = {
      "type": "bubble",
      "hero": {
        "type": "image",
        "url": "https://picsum.photos/1200/600?random=1",
        "size": "full",
        "aspectRatio": "20:9",
        "aspectMode": "cover"
      },
      "body": {
        "type": "box",
        "layout": "vertical",
        "spacing": "md",
        "contents": [
          {
            "type": "text",
            "text": "航班統計選單",
            "weight": "bold",
            "size": "xl"
          },
          {
            "type": "text",
            "text": "請選擇要下載的彙整表",
            "size": "sm",
            "color": "#888888"
          }
        ]
      },
      "footer": {
        "type": "box",
        "layout": "vertical",
        "spacing": "sm",
        "contents": [
          {
            "type": "button",
            "style": "primary",
            "height": "sm",
            "action": {
              "type": "postback",
              "label": "昨日航班彙整",
              "data": "action=yesterday_summary"
            }
          },
          {
            "type": "button",
            "style": "secondary",
            "height": "sm",
            "action": {
              "type": "postback",
              "label": "今日航班預估",
              "data": "action=today_forecast"
            }
          },
          {
            "type": "spacer",
            "size": "sm"
          }
        ],
        "flex": 0
      }
    }
    return FlexSendMessage(alt_text="航班統計選單", contents=flex)

@app.route("/healthz", methods=["GET"])
def healthz():
    return "ok"

@app.route("/callback", methods=["POST"])
def callback():
    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"

@handler.add(MessageEvent, message=TextMessage)
def handle_message(event: MessageEvent):
    text = (event.message.text or "").strip()

    # 關鍵字觸發選單（可自行擴充）
    if text in ["menu", "選單", "開始", "航班", "功能", "圖文選單"]:
        line_bot_api.reply_message(event.reply_token, build_menu_flex())
        return

    # 也支援直接輸入兩個功能名稱（展示用）
    if text in ["昨日航班彙整"]:
        url = "https://example.com/demo/yesterday_flight_summary.xlsx"
        msg = f"✅ 這是展示連結（假的）：\n{url}"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
        return

    if text in ["今日航班預估"]:
        url = "https://example.com/demo/today_flight_forecast.xlsx"
        msg = f"✅ 這是展示連結（假的）：\n{url}"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
        return

    # 其他文字 → 提示使用「選單」
    tip = "請輸入「選單」呼叫功能🙂"
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text=tip))

@handler.add(PostbackEvent)
def handle_postback(event: PostbackEvent):
    data = event.postback.data or ""
    if "action=yesterday_summary" in data:
        url = "https://example.com/demo/yesterday_flight_summary.xlsx"
        msg = f"🗓️ 昨日航班彙整（展示連結）\n{url}"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
        return

    if "action=today_forecast" in data:
        url = "https://example.com/demo/today_flight_forecast.xlsx"
        msg = f"📊 今日航班預估（展示連結）\n{url}"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
        return

    # 預設
    line_bot_api.reply_message(event.reply_token, TextSendMessage(text="已收到選擇。"))
