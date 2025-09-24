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
SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1Nttc45OMeYl5SysfxWJ0B5qUu9Bo42Hx/pub?gid=74488037&single=true&output=csv"  # 更新為你提供的工作表 gid=74488037  # 建議使用『發佈到網路』的公開 CSV 端點，較穩定  # 更新為正確的工作表 gid

# ---- 讀取 Google Sheets 並整理昨日資料（按固定儲存格） ----
def _col_letters_to_idx(col_letters: str) -> int:
    # A->0, B->1, ... Z->25, AA->26, ...
    col_letters = col_letters.strip().upper()
    n = 0
    for ch in col_letters:
        if not ('A' <= ch <= 'Z'):
            raise ValueError(f"Invalid column letter: {col_letters}")
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n - 1

def _read_cells_from_csv(url: str, refs: dict) -> dict:
    """refs: {key: (col_letters, row_number)} → return {key: value_str}
    會自動嘗試多種 CSV 端點，避免 400/403 問題。
    """
    import csv
    urls = [
        # 1) 發佈到網路（建議）：File → Share → Publish to web
        f"https://docs.google.com/spreadsheets/d/1Nttc45OMeYl5SysfxWJ0B5qUu9Bo42Hx/pub?gid=74488037&single=true&output=csv",
        # 2) 匯出端點（需檔案對外可檢視）：
        f"https://docs.google.com/spreadsheets/d/1Nttc45OMeYl5SysfxWJ0B5qUu9Bo42Hx/export?format=csv&gid=74488037",
        # 3) gviz 查詢端點（公開可檢視即可）：
        f"https://docs.google.com/spreadsheets/d/1Nttc45OMeYl5SysfxWJ0B5qUu9Bo42Hx/gviz/tq?tqx=out:csv&gid=74488037",
    ]

    last_err = None
    text = None
    for u in urls:
        try:
            resp = requests.get(u, headers={"User-Agent": "Mozilla/5.0"}, timeout=20, allow_redirects=True)
            resp.raise_for_status()
            # 有些情況會回 HTML（未公開或權限不足），簡單檢查一下
            if resp.text.strip().startswith("<!DOCTYPE html"):
                last_err = Exception("Got HTML instead of CSV (likely permission not public)")
                continue
            text = resp.text
            break
        except Exception as e:
            last_err = e
            continue

    if text is None:
        raise last_err or Exception("Failed to fetch CSV")

    rows = list(csv.reader(text.splitlines()))
    out = {}
    for key, (col_letters, row_num) in refs.items():
        r = int(row_num) - 1  # 1-based to 0-based
        c = _col_letters_to_idx(col_letters)
        try:
            out[key] = rows[r][c].strip()
        except Exception:
            out[key] = "-"
    return out

def fetch_yesterday_summary():
    try:
        today = datetime.date.today()
        yesterday = today - datetime.timedelta(days=1)
        target_date = yesterday.strftime("%Y/%m/%d")

        refs = {
            "kinmen_flights": ("CP", 8),
            "kinmen_seats": ("CQ", 8),
            "kinmen_pax": ("CR", 8),
            "kinmen_load": ("CS", 8),

            "penghu_flights": ("CP", 13),
            "penghu_seats": ("CQ", 13),
            "penghu_pax": ("CR", 13),
            "penghu_load": ("CS", 13),

            "matsu_flights": ("CP", 19),
            "matsu_seats": ("CQ", 19),
            "matsu_pax": ("CR", 19),
            "matsu_load": ("CS", 19),

            "main_flights": ("CP", 24),
            "main_seats": ("CQ", 24),
            "main_pax": ("CR", 24),
            "main_load": ("CS", 24),

            "other_flights": ("CP", 31),
            "other_seats": ("CQ", 31),
            "other_pax": ("CR", 31),
            "other_load": ("CS", 31),
        }

        cells = _read_cells_from_csv(SHEET_CSV_URL, refs)

        parts = [f"昨日({target_date})國內線簡要統計"]
        parts.append(
            f"\n金門航線：\n"
            f"✈️ 架次：{cells['kinmen_flights']}\n"
            f"💺 座位數：{cells['kinmen_seats']}\n"
            f"👥 載客數：{cells['kinmen_pax']}\n"
            f"📊 載客率：{cells['kinmen_load']}"
        )
        parts.append(
            f"\n澎湖航線：\n"
            f"✈️ 架次：{cells['penghu_flights']}\n"
            f"💺 座位數：{cells['penghu_seats']}\n"
            f"👥 載客數：{cells['penghu_pax']}\n"
            f"📊 載客率：{cells['penghu_load']}"
        )
        parts.append(
            f"\n馬祖航線：\n"
            f"✈️ 架次：{cells['matsu_flights']}\n"
            f"💺 座位數：{cells['matsu_seats']}\n"
            f"👥 載客數：{cells['matsu_pax']}\n"
            f"📊 載客率：{cells['matsu_load']}"
        )
        parts.append(
            f"\n本島航線：\n"
            f"✈️ 架次：{cells['main_flights']}\n"
            f"💺 座位數：{cells['main_seats']}\n"
            f"👥 載客數：{cells['main_pax']}\n"
            f"📊 載客率：{cells['main_load']}"
        )
        parts.append(
            f"\n其他離島航線：\n"
            f"✈️ 架次：{cells['other_flights']}\n"
            f"💺 座位數：{cells['other_seats']}\n"
            f"👥 載客數：{cells['other_pax']}\n"
            f"📊 載客率：{cells['other_load']}"
        )

        return "\n".join(parts)

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
