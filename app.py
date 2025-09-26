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
# 常數：Google Sheets 與對照
# =========================
WEEKLY_FILE_ID = "1Nttc45OMeYl5SysfxWJ0B5qUu9Bo42Hx"
WEEKLY_CSV_URL = (
    f"https://docs.google.com/spreadsheets/d/{WEEKLY_FILE_ID}/gviz/tq?"
    "tqx=out:csv&sheet=%E7%B5%B1%E8%A8%881&range=B1:DE32"
)

WEEKLY_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/1Nttc45OMeYl5SysfxWJ0B5qUu9Bo42Hx/edit?usp=drive_link&ouid=104418630202835382297&rtpof=true&sd=true"
)
# 列號為 1-based（你提供的 mapping）
ROW_MAP = {
    "各航線摘要統計": 31,  # 全航線彙總卡
    "金門航線": 7,
    "澎湖航線": 13,
    "馬祖航線": 18,
    "本島航線": 23,
    "其他離島航線": 30,
}
ROUTE_ORDER = [
    "各航線摘要統計",
    "金門航線",
    "澎湖航線",
    "馬祖航線",
    "本島航線",
    "其他離島航線",
]

DAILY_FILE_ID = "1KTPwIgiqB2AOoQI4P_TySam0l12DO7wd"
DAILY_CSV_URL = (
    f"https://docs.google.com/spreadsheets/d/{DAILY_FILE_ID}/gviz/tq?"
    "tqx=out:csv&sheet=%E5%9C%8B%E5%85%A7%E7%B7%9A&range=D1:P38"
)
CELL_SCHEDULED = "M19"
CELL_FLOWN = "M34"
CELL_CANCELLED = "M28"
DAILY_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/1KTPwIgiqB2AOoQI4P_TySam0l12DO7wd/edit?usp=drive_link&ouid=104418630202835382297&rtpof=true&sd=true"
)

# --- Daily 路線對照（以「擷取後的 CSV A1 座標」為準）---
DAILY_CANCEL_MAP: Dict[str, str] = {
    "金門航線": "C28",
    "澎湖航線": "F28",
    "馬祖航線": "I28",
    "花蓮航線": "J28",
    "臺東航線": "K28",
    "其他航線": "L28",
}

DAILY_FLOWN_MAP: Dict[str, Tuple[str, str]] = {
    "金門航線": ("C34", "C19"),
    "澎湖航線": ("F34", "F19"),
    "馬祖航線": ("I34", "I19"),
    "花蓮航線": ("J34", "J19"),
    "臺東航線": ("K34", "K19"),
    "其他航線": ("L34", "L19"),
}

# =========================
# 抽取器（Extractor）：只做資料萃取，回傳簡單 dict
# =========================

def extract_weekly(rows: List[List[str]]) -> Dict[str, Any]:
    def row_vals(row_1_based: int) -> Tuple[str, str, str, str]:
        i = row_1_based - 1
        if i < 0 or i >= len(rows):
            return ("-", "-", "-", "-")
        r = rows[i]
        get = lambda j: (r[j].strip() if j < len(r) and r[j] is not None else "-")
        return (get(0), get(1), get(2), get(3))  # CP, CQ, CR, CS

    dates = date_pack_for_ui()
    data = {
        "cover": {"start": dates["start7"], "end": dates["today"]},
        "yesterday": dates["yesterday"],
        "routes": []
    }
    for title in ROUTE_ORDER:
        row_index = ROW_MAP[title]
        cp, cq, cr, cs = row_vals(row_index)
        data["routes"].append({
            "title": title,
            "cp": cp, "cq": cq, "cr": cr, "cs": cs,
        })
    return data
  
def weekly_apply_v2(rows: List[List[str]], data: Dict[str, Any]) -> Dict[str, Any]:
    """套用新版 weekly 規格：
    - 日期從 CG2（MM月DD日，=昨日）換算成 YYYY/MM/DD，並回推 7 日。
    - 各卡資料來源欄位更新為指定的 A1（依序：架次、座位數、載客數、載客率）。
    """
    import re
    def _parse_mmdd_zh_to_date(mmdd_text: str) -> datetime.date:
        m = re.search(r"(\d{1,2})月(\d{1,2})日", mmdd_text or "")
        today = now_tw().date()
        if not m:
            return today - datetime.timedelta(days=1)
        y = today.year
        mm = int(m.group(1)); dd = int(m.group(2))
        try:
            d = datetime.date(y, mm, dd)
        except Exception:
            return today - datetime.timedelta(days=1)
        # 跨年處理：若推得日期晚於今天，表示實際是去年
        if d > today:
            d = datetime.date(y - 1, mm, dd)
        return d

    def _fmt(d: datetime.date) -> str:
        return d.strftime("%Y/%m/%d")

    # 1) 由 CG2 算日期區間
    cg2_text = get_a1(rows, "CG2", "")
    end_date = _parse_mmdd_zh_to_date(cg2_text)
    start_date = end_date - datetime.timedelta(days=7)
    data["cover"] = {"start": _fmt(start_date), "end": _fmt(end_date)}
    data["yesterday"] = _fmt(end_date)

    # 2) 各卡資料來源欄位（依序：架次、座位數、載客數、載客率）
    cell_map: Dict[str, Tuple[str, str, str, str]] = {
        "各航線摘要統計": ("CO32", "CP32", "CQ32", "CR32"),
        "金門航線": ("CO8", "CP8", "CQ8", "CR8"),
        "澎湖航線": ("CO14", "CP14", "CQ14", "CR14"),
        "馬祖航線": ("CO19", "CP19", "CQ19", "CR19"),
        "本島航線": ("CO24", "CP24", "CQ24", "CR24"),
        "其他離島航線": ("CO31", "CP31", "CQ31", "CR31"),
    }

    for r in data.get("routes", []):
        title = r.get("title", "")
        if title in cell_map:
            c1, c2, c3, c4 = cell_map[title]
            r["cp"] = get_a1(rows, c1, r.get("cp", "-"))
            r["cq"] = get_a1(rows, c2, r.get("cq", "-"))
            r["cr"] = get_a1(rows, c3, r.get("cr", "-"))
            r["cs"] = get_a1(rows, c4, r.get("cs", "-"))
    return data

def extract_daily(rows: List[List[str]]) -> Dict[str, Any]:
    """
    1) 日期：抓擷取後的 A1 前 10 個字元（YYYY-MM-DD）。
    2) 其他數值：依固定儲存格（M19、M34、M28）。
    3) 新增：路線別取消摘要（C/F/I/J/K/L28），與路線別已飛摘要（34 與 19 列）。
    備註：CSV 範圍保持 D1:P38，日期已在擷取後的 A1。
    """
    def _to_int(x: str) -> int:
        try:
            return int(str(x).replace(',', '').strip())
        except Exception:
            return 0

    a1_raw = get_a1(rows, "A1", "-")
    report_date = a1_raw[:10] if a1_raw and len(a1_raw) >= 10 else now_tw().strftime("%Y-%m-%d")

    # 取消摘要（只保留 >0）
    cancel_routes = []
    for name, cell in DAILY_CANCEL_MAP.items():
        v = _to_int(get_a1(rows, cell, "0"))
        if v > 0:
            cancel_routes.append({"name": name, "count": v})

    # 已飛摘要（皆顯示，左值/右值）
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
# Renderer：只負責把 dict → Flex JSON（不碰資料來源）
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
    # 需求：第二張至第七張圖卡副標題統一為『昨日(YYYY/MM/DD)摘要統計』
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


def bubble_daily_cancel(date_str: str, items: List[Dict[str, Any]]) -> Dict[str, Any]:
    # 當日取消摘要（只顯示 >0 的路線；數字紅色；數字緊接在標題後）
    list_contents: List[Dict[str, Any]] = []
    for it in items:
        list_contents.append({
            "type": "text",
            "text": f"{it['name']}：{it['count']}",
            "size": "lg",
            "contents": [
                {"type": "span", "text": f"{it['name']}："},
                {"type": "span", "text": str(it['count']), "weight": "bold", "color": "#C62828"}
            ]
        })
    if not list_contents:
        list_contents.append({"type": "text", "text": "（本日無取消）", "size": "sm", "color": "#888888"})

    return {
        "type": "bubble",
        "size": "mega",
        "body": {
            "type": "box",
            "layout": "vertical",
            "spacing": "md",
            "contents": [
                {"type": "text", "text": "當日取消摘要", "weight": "bold", "size": "lg"},
                {"type": "text", "text": f"{date_str}", "size": "sm", "color": "#888888"},
                {"type": "separator", "margin": "md"},
                {"type": "box", "layout": "vertical", "spacing": "sm", "margin": "md", "contents": list_contents}
            ]
        }
    }


def bubble_daily_flown(date_str: str, items: List[Dict[str, Any]]) -> Dict[str, Any]:
    # 當日已飛摘要（左值綠色 / 右值黑色；數字緊接在標題後）
    list_contents: List[Dict[str, Any]] = []
    for it in items:
        list_contents.append({
            "type": "text",
            "text": f"{it['name']}：{it['n1']}/{it['n2']}",
            "size": "lg",
            "contents": [
                {"type": "span", "text": f"{it['name']}："},
                {"type": "span", "text": str(it['n1']), "weight": "bold", "color": "#2E7D32"},
                {"type": "span", "text": "/", "weight": "bold"},
                {"type": "span", "text": str(it['n2']), "weight": "bold", "color": "#111111"}
            ]
        })

    return {
        "type": "bubble",
        "size": "mega",
        "body": {
            "type": "box",
            "layout": "vertical",
            "spacing": "md",
            "contents": [
                {"type": "text", "text": "當日已飛摘要", "weight": "bold", "size": "lg"},
                {"type": "text", "text": f"{date_str}", "size": "sm", "color": "#888888"},
                {"type": "separator", "margin": "md"},
                {"type": "box", "layout": "vertical", "spacing": "sm", "margin": "md", "contents": list_contents}
            ]
        }
    }


def flex_daily_payload(data: Dict[str, Any]) -> FlexSendMessage:
    def to_int(x):
        try:
            return int(str(x).replace(',', '').strip())
        except Exception:
            return None

    def pct(n, d):
        if n is None or d is None or d <= 0:
            return 0
        v = max(0, min(100, round(n * 100 / d)))
        return v

    sched_i = to_int(data["scheduled"])
    flown_i = to_int(data["flown"])
    canc_i = to_int(data["cancelled"])
    flown_pct = pct(flown_i, sched_i)
    cancel_pct = pct(canc_i, sched_i)

    # 第一頁：總覽
    bubble_overview = {
        "type": "bubble",
        "size": "mega",
        "body": {
            "type": "box",
            "layout": "vertical",
            "spacing": "md",
            "contents": [
                {"type": "text", "text": "國內線當日運量統計", "weight": "bold", "size": "lg"},
                {"type": "text", "text": f"{data['date']}摘要", "size": "sm", "color": "#888888"},
                {"type": "separator", "margin": "md"},
                {"type": "box", "layout": "horizontal", "margin": "md", "contents": [
                    {"type": "text", "text": "本日預計架次", "size": "lg", "color": "#333333", "flex": 2},
                    {"type": "box", "layout": "vertical", "flex": 3, "contents": [
                        {"type": "text", "text": str(data['scheduled'] or '-') , "size": "xxl", "weight": "bold", "color": "#111111", "align": "end"}
                    ], "alignItems": "flex-end"}
                ]},
                {"type": "box", "layout": "horizontal", "margin": "md", "contents": [
                    {"type": "text", "text": "已飛架次", "size": "lg", "color": "#2E7D32", "flex": 2},
                    {"type": "box", "layout": "vertical", "flex": 3, "contents": [
                        {"type": "text", "text": str(data['flown'] or '-') , "size": "xxl", "weight": "bold", "color": "#2E7D32", "align": "end"},
                        {"type": "text", "text": f"({flown_pct}%)", "size": "xs", "color": "#2E7D32", "align": "end"}
                    ], "alignItems": "flex-end"}
                ]},
                {"type": "box", "layout": "horizontal", "margin": "md", "contents": [
                    {"type": "text", "text": "取消架次", "size": "lg", "color": "#C62828", "flex": 2},
                    {"type": "box", "layout": "vertical", "flex": 3, "contents": [
                        {"type": "text", "text": str(data['cancelled'] or '-') , "size": "xxl", "weight": "bold", "color": "#C62828", "align": "end"},
                        {"type": "text", "text": f"({cancel_pct}%)", "size": "xs", "color": "#C62828", "align": "end"}
                    ], "alignItems": "flex-end"}
                ]},
                {"type": "button", "style": "link", "height": "sm", "action": {"type": "uri", "label": "開啟報表", "uri": data['sheet_url']}, "margin": "md"}
            ]
        },
        "styles": {"body": {"backgroundColor": "#FFFFFF"}}
    }

    # 第二頁：取消摘要
    bubble_cancel = bubble_daily_cancel(data['date'], data.get('cancel_routes', []))

    # 第三頁：已飛摘要
    bubble_flown = bubble_daily_flown(data['date'], data.get('flown_routes', []))

    return FlexSendMessage(
        alt_text=f"國內線當日運量統計（{data['date']}）",
        contents={"type": "carousel", "contents": [bubble_overview, bubble_cancel, bubble_flown]}
    )

# =========================
# Pipeline：從來源 → 抽取 → 渲染（兩個指令共用同一種流程）
# =========================

def build_weekly_flex_message() -> FlexSendMessage:
    rows = fetch_gviz_csv(WEEKLY_CSV_URL)
    data = extract_weekly(rows)          # 沿用原抽取
    data = weekly_apply_v2(rows, data)   # 套用新版欄位與日期規格
    return flex_weekly_payload(data)

def build_daily_flex_message() -> FlexSendMessage:
    rows = fetch_gviz_csv(DAILY_CSV_URL)
    data = extract_daily(rows)
    return flex_daily_payload(data)

# =========================
# 健康檢查（Cloud Run 健檢用）
@app.route("/healthz", methods=["GET", "HEAD"])
def healthz():
    # Cloud Run 偶爾會用 HEAD 做健檢；一律回 200
    if request.method == "HEAD":
        return "", 200
    return "ok", 200

# 額外提供根路由與 ready 檢查，避免 404
@app.route("/", methods=["GET"])
def root():
    return "ok", 200

@app.route("/readyz", methods=["GET", "HEAD"])
def readyz():
    if request.method == "HEAD":
        return "", 200
    return "ok", 200

# LINE Webhook
# =========================
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


if handler:
    @handler.add(MessageEvent, message=TextMessage)
    def handle_message(event: MessageEvent):
        text = (event.message.text or "").strip()

        if text == "7日內國內線統計表":
            try:
                line_bot_api.reply_message(event.reply_token, build_weekly_flex_message())
            except Exception as e:
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"7日內國內線統計表暫時無法使用 ({e})"))
            return

        if text == "國內線當日運量統計":
            try:
                line_bot_api.reply_message(event.reply_token, build_daily_flex_message())
            except Exception as e:
                today = date_pack_for_ui()["today"]
                line_bot_api.reply_message(event.reply_token, TextSendMessage(text=f"國內線當日運量統計 {today} 失敗 ({e})"))
            return

        # 其他文字先不回覆，避免干擾既有流程
        return


# 本機啟動（Cloud Run 可忽略）
if __name__ == "__main__":
    port = int(os.getenv("PORT", 8080))
    app.run(host="0.0.0.0", port=port)


