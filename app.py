# app.py
# LINE Bot：使用者輸入「七日內國內線統計報表」→ 回覆指定的 Google Sheets 連結
# 保留健康檢查與版本端點方便 Cloud Run / K8s 使用

import os
import time
import socket
import platform
from datetime import datetime, timezone
from flask import Flask, request, abort, jsonify
import logging
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import (
    MessageEvent, TextMessage, TextSendMessage
)

app = Flask(__name__)
logging.basicConfig(level=logging.INFO)
logger = app.logger

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
@app.route("/healthz", methods=["GET"])
def healthz():
    payload = base_health_payload("ok")
    return jsonify(payload), 200

@app.route("/readyz", methods=["GET"])
def readyz():
    checks = {
        "env.LINE_CHANNEL_SECRET": bool(CHANNEL_SECRET),
        "env.LINE_CHANNEL_ACCESS_TOKEN": bool(CHANNEL_ACCESS_TOKEN),
    }
    is_ready = all(checks.values())
    payload = base_health_payload("ok" if is_ready else "fail")
    payload.update({"checks": checks})
    return jsonify(payload), 200 if is_ready else 503

@app.route("/version", methods=["GET"])
def version():
    return jsonify(base_health_payload("ok")), 200

# ---- LINE Webhook ----
@app.route("/callback", methods=["POST"])
def callback():
    if not handler:
        logger.error("/callback called but LINE handler not configured (missing env?)")
        return jsonify({"error": "LINE handler not configured"}), 503

    signature = request.headers.get("X-Line-Signature", "")
    body = request.get_data(as_text=True)
    logger.info("/callback received body length=%s", len(body))
    try:
        handler.handle(body, signature)
    except InvalidSignatureError:
        logger.exception("InvalidSignatureError - check CHANNEL_SECRET matches the channel")
        abort(400)
    except Exception:
        logger.exception("Unhandled error while handling webhook")
        abort(500)
    return "OK"

# ---- 訊息處理 ----
if handler:
    @handler.add(MessageEvent, message=TextMessage)
    def handle_message(event: MessageEvent):
        text = (event.message.text or "").strip()
        logger.info("Received message: %r", text)

        if text == "七日內國內線統計報表":
            url = "https://docs.google.com/spreadsheets/d/1Nttc45OMeYl5SysfxWJ0B5qUu9Bo42Hx/edit?usp=drive_link&ouid=104418630202835382297&rtpof=true&sd=true"
            msg = f"📈 七日內國內線統計報表：\n{url}"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
            logger.info("Replied with 7-day domestic link")
            return

        tip = "請輸入「七日內國內線統計報表」🙂"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=tip))
        logger.info("Replied with tip")

# ---- 根路由與模擬測試 ----
@app.route("/", methods=["GET"])
def index():
    return ("Flight Bot online. Try /healthz /readyz /version /simulate?q=七日內國內線統計報表", 200)

@app.route("/simulate", methods=["GET"])
def simulate():
    q = (request.args.get("q") or "").strip()
    if q == "七日內國內線統計報表":
        return jsonify({"reply": "📈 七日內國內線統計報表：https://docs.google.com/spreadsheets/d/1Nttc45OMeYl5SysfxWJ0B5qUu9Bo42Hx/edit?usp=drive_link&ouid=104418630202835382297&rtpof=true&sd=true"}), 200
    return jsonify({"reply": "請輸入「七日內國內線統計報表」🙂"}), 200
