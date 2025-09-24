# app.py
# 簡易 LINE Bot：使用者輸入「昨日航班統計」或「今日航班預估」→ 回覆假的 Excel 下載網址
# 加入健康檢查端點：/healthz（存活）、/readyz（就緒，檢查必要環境變數）、/version（版本與執行狀態）

import os
import time
import socket
import platform
from datetime import datetime, timezone
from flask import Flask, request, abort, jsonify
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage
)

app = Flask(__name__)

# ---- 基本設定 ----
START_TIME = time.time()
APP_VERSION = os.environ.get("APP_VERSION", "0.1.0-demo")
CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET")
CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")

line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN) if CHANNEL_ACCESS_TOKEN else None
handler = WebhookHandler(CHANNEL_SECRET) if CHANNEL_SECRET else None

# ---- 健康檢查工具函式 ----
def uptime_seconds() -> int:
    return int(time.time() - START_TIME)

def base_health_payload(status: str):
    return {
        "status": status,
        "service": "flight-bot",
        "version": APP_VERSION,
        "hostname": socket.gethostname(),
        "python": platform.python_version(),
        "pid": os.getpid(),
        "started_at": datetime.fromtimestamp(START_TIME, tz=timezone.utc).isoformat(),
        "uptime_sec": uptime_seconds(),
    }

# ---- 健康檢查端點 ----
@app.route("/healthz", methods=["GET"])  # liveness：存活檢查（輕量、永遠應回 200）
def healthz():
    payload = base_health_payload("ok")
    return jsonify(payload), 200

@app.route("/readyz", methods=["GET"])  # readiness：就緒檢查（檢查必要設定是否到位）
def readyz():
    checks = {
        "env.LINE_CHANNEL_SECRET": bool(CHANNEL_SECRET),
        "env.LINE_CHANNEL_ACCESS_TOKEN": bool(CHANNEL_ACCESS_TOKEN),
    }
    is_ready = all(checks.values())
    payload = base_health_payload("ok" if is_ready else "fail")
    payload.update({
        "checks": checks
    })
    return jsonify(payload), 200 if is_ready else 503

@app.route("/version", methods=["GET"])  # 提供版本與基本狀態給監控或除錯
def version():
    return jsonify(base_health_payload("ok")), 200

# ---- LINE Webhook ----
@app.route("/callback", methods=["POST"])
def callback():
    if not handler:
        # 若未配置 handler，直接回 503 讓外部監控知悉未就緒
        return jsonify({"error": "LINE handler not configured"}), 503

    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        abort(400)
    return "OK"

# ---- 訊息處理 ----
if handler:  # 僅在 handler 存在時註冊事件處理，避免啟動期例外
    @handler.add(MessageEvent, message=TextMessage)
    def handle_message(event: MessageEvent):
        text = (event.message.text or "").strip()

        if text == "昨日航班統計":
            url = "https://example.com/demo/yesterday_flight_summary.xlsx"
            msg = f"✅ 這是展示連結（假的）：\n{url}"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
            return

        if text == "今日航班預估":
            url = "https://example.com/demo/today_flight_forecast.xlsx"
            msg = f"✅ 這是展示連結（假的）：\n{url}"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
            return

        # 其他輸入 → 提示訊息
        tip = "請輸入「昨日航班統計」或「今日航班預估」🙂"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=tip))

# ---- 選用：根路由回應 ----
@app.route("/", methods=["GET"])  # 方便人工快速確認服務有回應
def index():
    return (
        "Flight Bot online. Try /healthz /readyz /version", 200
    )
