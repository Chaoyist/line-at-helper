# app.py
# 目標：統一兩個指令的處理流程：
# [讀取 Google Sheet] → [抽取資料 extractor] → [渲染 Flex template]
# 方便後續維護與擴充（同一種 pipeline）。

import os
import csv
import requests
import datetime
from typing import List, Tuple, Dict, Any
from flask import Flask, request, abort
from linebot import LineBotApi, WebhookHandler
from linebot.exceptions import InvalidSignatureError
from linebot.models import MessageEvent, TextMessage, TextSendMessage, FlexSendMessage

try:
    from zoneinfo import ZoneInfo  # Python 3.9+
except Exception:
    ZoneInfo = None

app = Flask(__name__)

CHANNEL_SECRET = os.environ.get("LINE_CHANNEL_SECRET")
CHANNEL_ACCESS_TOKEN = os.environ.get("LINE_CHANNEL_ACCESS_TOKEN")

line_bot_api = LineBotApi(CHANNEL_ACCESS_TOKEN) if CHANNEL_ACCESS_TOKEN else None
handler = WebhookHandler(CHANNEL_SECRET) if CHANNEL_SECRET else None

HTTP_HEADERS = {"User-Agent": "Mozilla/5.0 (FlightBot)"}
HTTP_TIMEOUT = 20

# =========================
# 共用：時間與 Google Sheet
# =========================

def now_tw() -> datetime.datetime:
    try:
        return datetime.datetime.now(ZoneInfo("Asia/Taipei")) if ZoneInfo else datetime.datetime.now()
    except Exception:
        return datetime.datetime.now()


def date_pack_for_ui() -> Dict[str, str]:
    """提供 UI 會用到的日期字串：start/end/yesterday/today。"""
    today = now_tw()
    return {
        "today": today.strftime("%Y/%m/%d"),
        "yesterday": (today - datetime.timedelta(days=1)).strftime("%Y/%m/%d"),
        "start7": (today - datetime.timedelta(days=7)).strftime("%Y/%m/%d"),  # 不含今天共 7 天
    }


# --- 1~5分鐘快取設定 ---
CACHE_TTL_SECONDS = int(os.getenv("GVIZ_CACHE_TTL", "300"))  # 預設 300s，可用環境變數覆寫
GVIZ_CACHE: Dict[str, Tuple[float, List[List[str]]]] = {}

def fetch_gviz_csv(url: str) -> List[List[str]]:
    # 先讀快取
    try:
        exp_ts, cached_rows = GVIZ_CACHE.get(url, (0.0, None))  # type: ignore
        if cached_rows is not None and exp_ts > now_tw().timestamp():
            return cached_rows
    except Exception:
        pass

    resp = requests.get(url, headers=HTTP_HEADERS, timeout=HTTP_TIMEOUT)
    resp.raise_for_status()
    text = resp.text.strip()
    if text.startswith("<!DOCTYPE html"):
        raise RuntimeError("CSV endpoint returned HTML – check sharing/publish settings")
    rows = list(csv.reader(text.splitlines()))

    # 寫入快取
    try:
        GVIZ_CACHE[url] = (now_tw().timestamp() + CACHE_TTL_SECONDS, rows)
    except Exception:
        pass
    return rows


def a1_to_index(a1: str) -> Tuple[int, int]:
    s = a1.strip().upper()
    i = 0
    while i < len(s) and s[i].isalpha():
        i += 1
    col_letters, row_digits = s[:i], s[i:]
    if not col_letters or not row_digits.isdigit():
        raise ValueError(f"Invalid A1: {a1}")
    col_num = 0
    for ch in col_letters:
        col_num = col_num * 26 + (ord(ch) - ord('A') + 1)
    return (int(row_digits) - 1, col_num - 1)


def get_a1(rows: List[List[str]], a1: str, default: str = "-") -> str:
    r, c = a1_to_index(a1)
    if r < 0 or r >= len(rows):
        return default
    row = rows[r]
    if c < 0 or c >= len(row):
        return default
    return (row[c] or "").strip() or default

# =========================
# 常數：Google Sheets
# =========================
# Weekly A1 對應（依你提供的 CSV 轉換後座標，順序= 架次、座位數、載客數、載客率）
WEEKLY_routes: List[Dict[str, Any]] = []
    for title in ROUTE_ORDER:
        c1, c2, c3, c4 = WEEKLY_CELL_MAP[title]
        routes.append({
            "title": title,
            "cp": get_a1(rows, c1, "-"),
            "cq": get_a1(rows, c2, "-"),
            "cr": get_a1(rows, c3, "-"),
            "cs": get_a1(rows, c4, "-"),
        })

    return {
        "cover": {"start": _fmt(start_date), "end": _fmt(end_date)},
        "yesterday": _fmt(end_date),
        "routes": routes,
    }


def extract_daily(rows: List[List[str]]) -> Dict[str, Any]:
    """
    1) 日期：抓擷取後的 A1 前 10 個字元（YYYY-MM-DD）。
    2) 其他數值：依固定儲存格（M19、M34、M28）。
    3) 新增：路線別取消摘要與已飛摘要。
    """
    def _to_int(x: str) -> int:
        try:
            return int(str(x).replace(',', '').strip())
        except Exception:
            return 0

    a1_raw = get_a1(rows, "A1", "-")
    report_date = a1_raw[:10] if a1_raw and len(a1_raw) >= 10 else now_tw().strftime("%Y-%m-%d")

    cancel_routes = []
    for name, cell in DAILY_CANCEL_MAP.items():
        v = _to_int(get_a1(rows, cell, "0"))
        if v > 0:
            cancel_routes.append({"name": name, "count": v})

    flown_routes = []
    for name, (c1, c2) in DAILY_FLOWN_MAP.items():
        n1 = _to_int(get_a1(rows, c1, "0"))
        n2 = _to_int(get_a1(rows, c2, "0"))
        flown_routes.append({"name": name, "n1": n1, "n2": n2})

    return {
        "date": report_date,
        "scheduled": get_a1(rows, CELL_SCHEDULED, "-"),
        "flown": get_a1(rows, CELL_FLOWN, "-"),
        "cancelled": get_a1(rows, CELL_CANCELLED, "-"),
        "sheet_url": DAILY_SHEET_URL,
        "cancel_routes": cancel_routes,
        "flown_routes": flown_routes,
    }

# =========================
# Renderer
# =========================

def bubble_cover(start: str, end: str) -> Dict[str, Any]:
    return {
        "type": "bubble",
        "size": "mega",
        "body": {
            "type": "box",
            "layout": "vertical",
            "spacing": "md",
            "contents": [
                {"type": "text", "text": "7日內國內線統計表", "weight": "bold", "size": "lg"},
                {"type": "text", "text": f"{start}-{end}", "size": "sm", "color": "#888888"},
                {"type": "separator", "margin": "md"},
                {"type": "button", "style": "link", "height": "sm",
                 "action": {"type": "uri", "label": "開啟報表", "uri": WEEKLY_SHEET_URL}},
                {
                    "type": "box",
                    "layout": "horizontal",
                    "margin": "lg",
                    "justifyContent": "center",
                    "alignItems": "center",
                    "contents": [
                        {"type": "text", "text": "⬅️ 往左滑看昨日各航線摘要統計", "size": "xs", "color": "#666666", "alignItems": "center"}
                    ]
                }
            ]
        }
    }


def bubble_route(title: str, ymd_yesterday: str, cp: str, cq: str, cr: str, cs: str) -> Dict[str, Any]:
    subtitle = f"昨日({ymd_yesterday})摘要統計"
    return {
        "type": "bubble",
        "size": "mega",
        "body": {
            "type": "box",
            "layout": "vertical",
            "spacing": "md",
            "contents": [
                {"type": "text", "text": title, "weight": "bold", "size": "lg"},
                {"type": "text", "text": subtitle, "size": "sm", "color": "#888888"},
                {"type": "separator", "margin": "md"},
                {"type": "box", "layout": "vertical", "spacing": "sm", "margin": "md", "contents": [
                    {"type": "text", "text": f"✈️ 架次：{cp}", "size": "md", "weight": "bold", "wrap": True},
                    {"type": "text", "text": f"💺 座位數：{cq}", "size": "md", "weight": "bold", "wrap": True},
                    {"type": "text", "text": f"👥 載客數：{cr}", "size": "md", "weight": "bold", "wrap": True},
                    {"type": "text", "text": f"📊 載客率：{cs}", "size": "md", "weight": "bold", "wrap": True}
                ]}
            ]
        }
    }


def flex_weekly_payload(data: Dict[str, Any]) -> FlexSendMessage:
    bubbles = [bubble_cover(data["cover"]["start"], data["cover"]["end"])]
    y = data["yesterday"]
    for item in data["routes"]:
        bubbles.append(bubble_route(item["title"], y, item["cp"], item["cq"], item["cr"], item["cs"]))
    return FlexSendMessage(alt_text="7日內國內線統計表", contents={"type": "carousel", "contents": bubbles})


def flex_daily_payload(data: Dict[str, Any]) -> FlexSendMessage:
    def to_int(x):
        try:
            return int(str(x).replace(',', '').strip())
        except Exception:
            return None

    sched_i = to_int(data.get("scheduled"))
    flown_i = to_int(data.get("flown"))
    canc_i = to_int(data.get("cancelled"))

    def pct(n, d):
        if n is None or d is None or d <= 0:
            return 0
        return max(0, min(100, round(n * 100 / d)))

    flown_pct = pct(flown_i, sched_i)
    cancel_pct = pct(canc_i, sched_i)

    # ===== 第一張：總覽 =====
    bubble_overview = {
        "type": "bubble",
        "size": "mega",
        "body": {
            "type": "box",
            "layout": "vertical",
            "spacing": "md",
            "contents": [
                {"type": "text", "text": "國內線當日運量統計", "weight": "bold", "size": "lg"},
                {"type": "text", "text": f"日期：{data['date']}", "size": "sm", "color": "#888888"},
                {"type": "separator", "margin": "md"},
                {"type": "box", "layout": "vertical", "spacing": "sm", "margin": "md", "contents": [
                    {"type": "box", "layout": "horizontal", "contents": [
                        {"type": "text", "text": "預計架次", "flex": 2, "size": "lg", "weight": "bold", "color": "#000000"},
                        {"type": "text", "text": str(data.get("scheduled", "-")), "flex": 1, "size": "xl", "align": "end", "weight": "bold", "color": "#000000"}
                    ]},
                    {"type": "box", "layout": "vertical", "contents": [
                        {"type": "box", "layout": "horizontal", "contents": [
                            {"type": "text", "text": "已飛架次", "flex": 2, "size": "lg", "weight": "bold", "color": "#000000"},
                            {"type": "text", "text": str(data.get("flown", "-")), "flex": 1, "size": "xl", "align": "end", "weight": "bold", "color": "#16A34A"}
                        ]},
                        {"type": "text", "text": f"({flown_pct}%)", "size": "xs", "align": "end", "color": "#16A34A"}
                    ]},
                    {"type": "box", "layout": "vertical", "contents": [
                        {"type": "box", "layout": "horizontal", "contents": [
                            {"type": "text", "text": "取消架次", "flex": 2, "size": "lg", "weight": "bold", "color": "#000000"},
                            {"type": "text", "text": str(data.get("cancelled", "-")), "flex": 1, "size": "xl", "align": "end", "weight": "bold", "color": "#DC2626"}
                        ]},
                        {"type": "text", "text": f"({cancel_pct}%)", "size": "xs", "align": "end", "color": "#DC2626"}
                    ]}
                ]}
            ]
        }
    }

    bubbles = [bubble_overview]

    # ===== 第二張：當日取消摘要 =====
    if data.get("cancel_routes"):
        cancel_lines = []
        for x in data["cancel_routes"]:
            cancel_lines.append({
                "type": "box",
                "layout": "horizontal",
                "contents": [
                    {"type": "text", "text": f"{x['name']}：", "size": "lg", "wrap": False, "flex": 0},
                    {"type": "text", "text": str(x['count']), "size": "lg", "weight": "bold", "color": "#DC2626", "wrap": False}
                ]
            })
        bubbles.append({
            "type": "bubble",
            "size": "mega",
            "body": {
                "type": "box",
                "layout": "vertical",
                "spacing": "md",
                "contents": [
                    {"type": "text", "text": "當日取消摘要", "weight": "bold", "size": "lg"},
                    {"type": "separator", "margin": "md"},
                    {"type": "box", "layout": "vertical", "spacing": "sm", "contents": cancel_lines}
                ]
            }
        })

    # ===== 第三張：當日已飛摘要 =====
    if data.get("flown_routes"):
        flown_lines = []
        for x in data["flown_routes"]:
            # 將 57/80 連在一起顯示，57 綠色、/80 黑色，採用 span 分段著色
            value_text = {
                "type": "text",
                "size": "lg",
                "weight": "bold",
                "wrap": False,
                "contents": [
                    {"type": "span", "text": str(x['n1']), "color": "#16A34A"},
                    {"type": "span", "text": f"/{x['n2']}", "color": "#000000"}
                ]
            }
            flown_lines.append({
                "type": "box",
                "layout": "horizontal",
                "contents": [
                    {"type": "text", "text": f"{x['name']}：", "size": "lg", "wrap": False, "flex": 0},
                    value_text
                ]
            })
        bubbles.append({
            "type": "bubble",
            "size": "mega",
            "body": {
                "type": "box",
                "layout": "vertical",
                "spacing": "md",
                "contents": [
                    {"type": "text", "text": "當日已飛摘要", "weight": "bold", "size": "lg"},
                    {"type": "separator", "margin": "md"},
                    {"type": "box", "layout": "vertical", "spacing": "sm", "contents": flown_lines}
                ]
            }
        })

    return FlexSendMessage(alt_text="國內線當日運量統計", contents={"type": "carousel", "contents": bubbles})

# =========================
# Builder：把抽取與渲染串起來
# =========================

def build_weekly_flex_message() -> FlexSendMessage:
    rows = fetch_gviz_csv(WEEKLY_CSV_URL)
    data = extract_weekly(rows)          # 直接以 A1 座標抽取（含 CG2 日期區間）
    return flex_weekly_payload(data)


def build_daily_flex_message() -> FlexSendMessage:
    rows = fetch_gviz_csv(DAILY_CSV_URL)
    data = extract_daily(rows)
    return flex_daily_payload(data)

# =========================
# Flask 路由

@app.get("/weekly/debug")
def weekly_debug():
    """即時檢查 Weekly 每張卡實際抓到的儲存格與值，方便對版。
    會回傳各路線四個欄位（架次/座位/載客/載客率）的 A1 與值，以及 CG2 日期與 CSV 範圍。
    """
    rows = fetch_gviz_csv(WEEKLY_CSV_URL)
    out: Dict[str, Any] = {"_csv_range": "B1:DE32", "_yesterday_CG2": get_a1(rows, "CG2", "")}
    for title, (a, b, c, d) in WEEKLY_CELL_MAP.items():
        out[title] = {
            "架次": {"cell": a, "value": get_a1(rows, a, "")},
            "座位數": {"cell": b, "value": get_a1(rows, b, "")},
            "載客數": {"cell": c, "value": get_a1(rows, c, "")},
            "載客率": {"cell": d, "value": get_a1(rows, d, "")},
        }
    return out

# Flask 路由
# =========================

@app.get("/healthz")
def healthz():
    return {"status": "ok", "time": now_tw().isoformat()}


@app.post("/callback")
def callback():
    if not handler:
        return ("handler not configured", 500)

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
    reply: TextSendMessage | FlexSendMessage

    try:
        if text in ["7日內國內線統計表", "7日內統計", "7日統計", "7日內"]:
            reply = build_weekly_flex_message()
        elif text in ["國內線當日運量統計", "當日運量", "今日國內線"]:
            reply = build_daily_flex_message()
        else:
            tips = (
                "可用指令：\n"
                "・7日內國內線統計表\n"
                "・國內線當日運量統計"
            )
            reply = TextSendMessage(text=tips)
    except Exception as e:
        reply = TextSendMessage(text=f"查詢失敗：{e}\n請確認資料來源是否可讀或欄位是否異動。")

    if line_bot_api:
        line_bot_api.reply_message(event.reply_token, reply)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8080"))
    app.run(host="0.0.0.0", port=port)
