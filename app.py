# app.py
# LINE Bot：
# - 使用者輸入「7日內國內線統計表」：回覆短網址 + 摘要（來源：Google Sheets gviz CSV）。
# - 使用者輸入「國內線當日運量統計」：回覆短網址 + 本日三項統計（來源：Google Sheets gviz CSV）。：回覆短網址 + 本日三項統計（來源：Google Sheets gviz CSV）。
#
# 統一：
# 1) 以 gviz CSV 端點存取（免 OAuth，前提是表單已設「知道連結的人可檢視」）。
# 2) 以通用 fetch_gviz_csv(url) 取回二維陣列 rows。
# 3) 以 get_a1(rows, "M19") 讀取 A1 位置；以 get_row_values(rows, row_1_based, n) 讀整列前 n 欄。
# 4) 清楚命名與註解，便於後續維護與擴充。

import os
import csv
import requests
import datetime
from typing import List, Tuple
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage, FlexSendMessage

# 時區（Python 3.9+ 內建 zoneinfo），用於顯示台灣時間日期
try:
    from zoneinfo import ZoneInfo  # Python 3.9+
except Exception:
    ZoneInfo = None

# ---------------------------------
# Flask / LINE 基本設定
# ---------------------------------
app = Flask(__name__)

CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET")
CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")

line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN) if CHANNEL_ACCESS_TOKEN else None
handler = WebhookHandler(CHANNEL_SECRET) if CHANNEL_SECRET else None

# ---------------------------------
# Google Sheets gviz CSV 共同設定
# ---------------------------------
HTTP_HEADERS = {"User-Agent": "Mozilla/5.0 (FlightBot)"}
HTTP_TIMEOUT = 20

# ---- 時間工具：取得台灣今天日期字串（有 zoneinfo 用之，無則安全退回）----
def today_str_tw(fmt: str = "%Y/%m/%d") -> str:
    try:
        if ZoneInfo:
            return datetime.datetime.now(ZoneInfo("Asia/Taipei")).strftime(fmt)
    except Exception:
        # 有些精簡映像沒有 tzdata，這裡安全退回系統時間
        pass
    return datetime.datetime.now().strftime(fmt)

# ---- 7日內國內線統計表（固定分頁 + 範圍）----
WEEKLY_FILE_ID = "1Nttc45OMeYl5SysfxWJ0B5qUu9Bo42Hx"
WEEKLY_CSV_URL = (
    f"https://docs.google.com/spreadsheets/d/{WEEKLY_FILE_ID}/gviz/tq?"
    "tqx=out:csv&sheet=%E7%B5%B1%E8%A8%881&range=CP2:CS32"
)
# 對應列（1-based 索引，以你提供的行號）
ROW_MAP = {
    "全航線": 31,
    "金門航線": 7,
    "澎湖航線": 13,
    "馬祖航線": 18,
    "本島航線": 23,
    "其他離島航線": 30,
}
ROUTE_LIST = ["金門航線", "澎湖航線", "馬祖航線", "本島航線", "其他離島航線"]

# ---- 當日疏運統計表（國內線 D1:P38）----
DAILY_FILE_ID = "1KTPwIgiqB2AOoQI4P_TySam0l12DO7wd"
DAILY_CSV_URL = (
    f"https://docs.google.com/spreadsheets/d/{DAILY_FILE_ID}/gviz/tq?"
    "tqx=out:csv&sheet=%E5%9C%8B%E5%85%A7%E7%B7%9A&range=D1:P38"
)
# 需要的 A1 位置
CELL_SCHEDULED = "M19"  # 本日表定架次
CELL_FLOWN = "M34"      # 已飛架次
CELL_CANCELLED = "M28"  # 取消架次

# ---------------------------------
# 通用：下載 gviz CSV 並轉成 rows
# ---------------------------------
def fetch_gviz_csv(url: str) -> List[List[str]]:
    """下載 gviz CSV（免 OAuth）。回傳二維陣列 rows。出錯拋例外。"""
    resp = requests.get(url, headers=HTTP_HEADERS, timeout=HTTP_TIMEOUT)
    resp.raise_for_status()
    text = resp.text.strip()
    # 若回 HTML 多半表示權限或重導（非 out:csv）
    if text.startswith("<!DOCTYPE html"):
        raise RuntimeError("CSV endpoint returned HTML – check sharing/publish settings")
    rows = list(csv.reader(text.splitlines()))
    return rows

# ---------------------------------
# A1 工具：將 A1 轉 (row_idx, col_idx) 以及從 rows 取值
# ---------------------------------

def a1_to_index(a1: str) -> Tuple[int, int]:
    """A1 → 0-based (row_idx, col_idx)。例如 'M19' → (18, 12)。"""
    s = a1.strip().upper()
    i = 0
    while i < len(s) and s[i].isalpha():
        i += 1
    col_letters = s[:i]
    row_digits = s[i:]
    if not col_letters or not row_digits.isdigit():
        raise ValueError(f"Invalid A1: {a1}")
    # 欄位：A=1 → Z=26 → AA=27 → ...
    col_num = 0
    for ch in col_letters:
        col_num = col_num * 26 + (ord(ch) - ord('A') + 1)
    col_idx = col_num - 1
    row_idx = int(row_digits) - 1
    return (row_idx, col_idx)

def get_a1(rows: List[List[str]], a1: str, default: str = "-") -> str:
    r, c = a1_to_index(a1)
    if r < 0 or r >= len(rows):
        return default
    row = rows[r]
    if c < 0 or c >= len(row):
        return default
    return (row[c] or "").strip() or default

# ---------------------------------
# 功能一：7日內國內線統計表（回覆全航線 + 各航線四欄）
# ---------------------------------

def build_weekly_summary_text() -> str:
    """抓取 WEEKLY_CSV_URL，組成多段摘要文字。若失敗回錯誤說明。"""
    try:
        rows = fetch_gviz_csv(WEEKLY_CSV_URL)

        def get_row_values(row_1_based: int, n: int = 4) -> Tuple[str, ...]:
            i = row_1_based - 1
            if i < 0 or i >= len(rows):
                return tuple(["-"] * n)
            r = rows[i]
            vals = []
            for j in range(n):
                vals.append((r[j].strip() if j < len(r) and r[j] is not None else "-"))
            return tuple(vals)

        # 標題：昨日(YYYY/MM/DD)航班彙整摘要（台灣時間）
        # 以昨天日期呈現（使用共用工具，避免 tzdata 缺失崩潰）
        try:
            y = datetime.datetime.strptime(today_str_tw(), "%Y/%m/%d") - datetime.timedelta(days=1)
        except Exception:
            y = datetime.datetime.now() - datetime.timedelta(days=1)
        title = f"\n\n昨日({y.strftime('%Y/%m/%d')})航班彙整摘要"

        parts = []
        parts.append(title)
        # 全航線
        cp, cq, cr, cs = get_row_values(ROW_MAP["全航線"])  # CP=架次 CQ=座位 CR=載客 CS=載客率
        parts.append("全航線：")
        parts.append(f"✈️ 架次：{cp}")
        parts.append(f"💺 座位數：{cq}")
        parts.append(f"👥 載客數：{cr}")
        parts.append(f"📊 載客率：{cs}")
        # 各航線
        for route in ROUTE_LIST:
            cp, cq, cr, cs = get_row_values(ROW_MAP[route])
            parts.append(f"\n{route}：")
            parts.append(f"✈️ 架次：{cp}")
            parts.append(f"💺 座位數：{cq}")
            parts.append(f"👥 載客數：{cr}")
            parts.append(f"📊 載客率：{cs}")

        return "\n".join(parts)
    except Exception as e:
        return f"（暫時無法取得統計資料：{e}）"

# ---------------------------------
# 功能二：當日疏運統計表（三項 A1 欄位）
# ---------------------------------

def fetch_daily_transport_summary() -> Tuple[str, str, str]:
    """
    擷取「當日疏運統計表」摘要三值：
    本日表定架次=M19、已飛架次=M34、取消架次=M28。
    任何錯誤一律以 '-' 回傳避免中斷。
    """
    try:
        rows = fetch_gviz_csv(DAILY_CSV_URL)
        scheduled = get_a1(rows, CELL_SCHEDULED, "-")
        flown = get_a1(rows, CELL_FLOWN, "-")
        cancelled = get_a1(rows, CELL_CANCELLED, "-")
        return (scheduled, flown, cancelled)
    except Exception:
        return ("-", "-", "-")

# ---------------------------------
# Flex Message：主KPI卡（堆疊條：已飛 + 取消 = 表定）
# ---------------------------------

def build_daily_kpi_flex(scheduled: str, flown: str, cancelled: str, date_str: str, url: str) -> FlexSendMessage:
    """
    國內線當日運量統計（數字版）：
    - 標題：國內線當日運量統計
    - 副標：YYYY/MM/DD摘要
    - 本日預計架次（黑色大字，與下方對齊）
    - 已飛架次（綠色大字，括號顯示百分比，同一行）
    - 取消架次（紅色大字，括號顯示百分比，同一行）
    - 下方：開啟報表連結按鈕
    """
    def to_int(x):
        try:
            return int(str(x).replace(',', '').strip())
        except Exception:
            return None

    sched_i = to_int(scheduled)
    flown_i = to_int(flown)
    canc_i  = to_int(cancelled)

    def pct(n, d):
        if n is None or d is None or d <= 0:
            return 0
        v = max(0, min(100, round(n * 100 / d)))
        return v

    flown_pct = pct(flown_i, sched_i)
    cancel_pct = pct(canc_i, sched_i)

    s_scheduled = scheduled if scheduled else "-"
    s_flown     = flown if flown else "-"
    s_cancelled = cancelled if cancelled else "-"

    bubble = {
        "type": "bubble",
        "size": "mega",
        "body": {
            "type": "box",
            "layout": "vertical",
            "spacing": "md",
            "contents": [
                {"type": "text", "text": "國內線當日運量統計", "weight": "bold", "size": "lg"},
                {"type": "text", "text": f"{date_str}摘要", "size": "sm", "color": "#888888"},
                {"type": "separator", "margin": "md"},

                {"type": "box", "layout": "horizontal", "margin": "md", "contents": [
                    {"type": "text", "text": "本日預計架次", "size": "sm", "color": "#333333", "flex": 2},
                    {"type": "text", "text": str(s_scheduled), "size": "xxl", "weight": "bold", "align": "end", "flex": 3, "color": "#111111"}
                ]},

                {"type": "box", "layout": "horizontal", "margin": "md", "contents": [
                    {"type": "text", "text": "已飛架次", "size": "sm", "color": "#2E7D32", "flex": 2},
                    {"type": "text", "text": f"{s_flown} ({flown_pct}%)", "size": "xxl", "weight": "bold", "color": "#2E7D32", "align": "end", "flex": 3}
                ]},

                {"type": "box", "layout": "horizontal", "margin": "md", "contents": [
                    {"type": "text", "text": "取消架次", "size": "sm", "color": "#C62828", "flex": 2},
                    {"type": "text", "text": f"{s_cancelled} ({cancel_pct}%)", "size": "xxl", "weight": "bold", "color": "#C62828", "align": "end", "flex": 3}
                ]},

                {"type": "button", "style": "link", "height": "sm", "action": {"type": "uri", "label": "開啟報表", "uri": url}, "margin": "md"}
            ]
        },
        "styles": {"body": {"backgroundColor": "#FFFFFF"}}
    }

    return FlexSendMessage(alt_text=f"國內線當日運量統計（{date_str}）", contents=bubble)

# ---------------------------------
# LINE Webhook / 路由
# ---------------------------------
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
            summary = build_weekly_summary_text()
            msg = f"📈 7日內國內線統計表：\n{url}{summary and ('' + summary)}"
            line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
            return

        if text == "國內線當日運量統計":
            url = "https://docs.google.com/spreadsheets/d/1KTPwIgiqB2AOoQI4P_TySam0l12DO7wd/edit?usp=drive_link&ouid=104418630202835382297&rtpof=true&sd=true"
            scheduled, flown, cancelled = fetch_daily_transport_summary()
            today = today_str_tw()
            try:
                flex = build_daily_kpi_flex(scheduled, flown, cancelled, today, url)
                line_bot_api.reply_message(event.reply_token, flex)
            except Exception as e:
                # 失敗退回文字版，附上DEBUG訊息
                msg = (
                    f"國內線當日運量統計{url}"
                    f"{today}摘要"
                    f"本日預計架次：{scheduled}"
                    f"已飛架次：{flown}"
                    f"取消架次：{cancelled}"
                    f"(DEBUG: {e})"
                )
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text=msg))
            return

        tip = "請輸入「7日內國內線統計表」或「國內線當日運量統計」🙂"
        line_bot_api.reply_message(event.reply_token, TextSendMessage(text=tip))

# ---- 根路由（健康檢查） ----
@app.route("/", methods=["GET"])
def index():
    return ("Flight Bot online. POST to /callback with LINE events", 200)
