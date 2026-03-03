#!/usr/bin/env python3
"""
PJ PL管理スプレッドシート — 整理版を新規作成
"""
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

TOKEN = "/Users/shinsukesaito/claude/ai-cost-monitor/token.json"
OLD_SSID = "1JWP6-2SuzTU5EkOQHR4hg2eBwGA3AfLTgIonIQ_CotE"

creds = Credentials.from_authorized_user_file(TOKEN)
sheets_svc = build("sheets", "v4", credentials=creds)
drive_svc = build("drive", "v3", credentials=creds)

# ─── 色定義 ───
WHITE = {"red": 1, "green": 1, "blue": 1}
DARK_BLUE = {"red": 0.15, "green": 0.25, "blue": 0.45}
LIGHT_BLUE = {"red": 0.85, "green": 0.92, "blue": 1.0}
LIGHT_GRAY = {"red": 0.95, "green": 0.95, "blue": 0.95}
LIGHT_GREEN = {"red": 0.85, "green": 0.95, "blue": 0.85}
LIGHT_RED = {"red": 1.0, "green": 0.9, "blue": 0.9}
LIGHT_YELLOW = {"red": 1.0, "green": 0.98, "blue": 0.85}
MID_BLUE = {"red": 0.6, "green": 0.75, "blue": 0.95}

# ─── PJ台帳データ ───
PJ_DATA = [
    {"code": "PJ0001", "id": "346", "name": "TMC（WBT）", "company": "ウーブン・バイ・トヨタ株式会社", "project": "オートレース日程システム", "type": "コンサル", "amount": 9475000, "start": "2026/01/05", "end": "2026/03/13", "months": 3, "monthly": 3158333, "sheet": "WBT"},
    {"code": "PJ0002", "id": "977", "name": "ASAP", "company": "ASAP SECURITY 株式会社", "project": "シフト管理システム", "type": "システム開発", "amount": 2800000, "start": "2025/11/01", "end": "2026/02/28", "months": 4, "monthly": 700000, "sheet": "ASAP"},
    {"code": "PJ0003", "id": "", "name": "タマチャンショップ", "company": "", "project": "", "type": "", "amount": 0, "start": "", "end": "", "months": 0, "monthly": 0, "sheet": None},
    {"code": "PJ0004", "id": "975", "name": "ユーグレナ", "company": "株式会社ユーグレナ", "project": "", "type": "提携", "amount": 0, "start": "", "end": "", "months": 1, "monthly": 0, "sheet": "ユーグレナ"},
    {"code": "PJ0005", "id": "979", "name": "清田軌道", "company": "清田軌道工業（株）", "project": "", "type": "研修", "amount": 4550000, "start": "2026/01/27", "end": "2026/01/28", "months": 1, "monthly": 4550000, "sheet": "清田軌道"},
    {"code": "PJ0006", "id": "970", "name": "ケイズ", "company": "株式会社ケイズグループ", "project": "", "type": "その他", "amount": 3500000, "start": "2025/10/01", "end": "2026/04/30", "months": 7, "monthly": 500000, "sheet": "ケイズ"},
    {"code": "PJ0007", "id": "988", "name": "栗田工業（IST）", "company": "栗田工業株式会社", "project": "IST案件", "type": "研修", "amount": 3900000, "start": "2025/11/01", "end": "2026/02/28", "months": 4, "monthly": 975000, "sheet": None},
    {"code": "PJ0008", "id": "207", "name": "小野薬品（AIMG）", "company": "小野薬品工業 株式会社", "project": "応用研修（AIMG）", "type": "研修", "amount": 1200000, "start": "2026/02/01", "end": "2026/02/28", "months": 1, "monthly": 1200000, "sheet": None},
    {"code": "PJ0009", "id": "304", "name": "原子力エンジニアリング", "company": "株式会社原子力エンジニアリング", "project": "研修＋サーベイ", "type": "研修", "amount": 700000, "start": "2026/02/02", "end": "2026/02/28", "months": 1, "monthly": 700000, "sheet": None},
    {"code": "PJ0010", "id": "346", "name": "WBT（重複）", "company": "ウーブン・バイ・トヨタ株式会社", "project": "オートレース日程システム", "type": "コンサル", "amount": 9475000, "start": "2026/01/05", "end": "2026/03/13", "months": 3, "monthly": 3158333, "sheet": None},
    {"code": "PJ0011", "id": "536", "name": "パナソニック", "company": "パナソニック株式会社 くらしアプライアンス社", "project": "", "type": "研修", "amount": 3000000, "start": "2026/02/01", "end": "2026/03/31", "months": 2, "monthly": 1500000, "sheet": None},
    {"code": "PJ0012", "id": "996", "name": "日本経営合理化協会", "company": "株式会社日本経営合理化協会事業団", "project": "", "type": "提携", "amount": 300000, "start": "2026/03/01", "end": "2026/03/17", "months": 1, "monthly": 300000, "sheet": None},
    {"code": "PJ0013", "id": "997", "name": "A.モンライン①", "company": "株式会社A.モンライン", "project": "AI ネイティブリーダーズプログラム(90日)", "type": "研修", "amount": 600000, "start": "2025/11/24", "end": "2026/02/23", "months": 3, "monthly": 200000, "sheet": None},
    {"code": "PJ0014", "id": "998", "name": "A.モンライン②", "company": "株式会社A.モンライン", "project": "経営幹部向け AI活用合宿", "type": "研修", "amount": 1500000, "start": "2025/11/01", "end": "2025/12/31", "months": 2, "monthly": 750000, "sheet": None},
    {"code": "PJ0015", "id": "1003", "name": "コレックHD", "company": "株式会社コレックホールディングス", "project": "", "type": "提携", "amount": 0, "start": "", "end": "", "months": 1, "monthly": 0, "sheet": None},
    {"code": "PJ0016", "id": "1025", "name": "有希化学", "company": "有希化学", "project": "GPT代行", "type": "その他", "amount": 600000, "start": "2026/01/01", "end": "2026/12/31", "months": 12, "monthly": 50000, "sheet": None},
    {"code": "PJ0017", "id": "1048", "name": "ONE COMPATH", "company": "株式会社ONE COMPATH", "project": "Shufoo! AI 2026年1-3月", "type": "システム開発", "amount": 3000000, "start": "2026/01/01", "end": "2026/03/31", "months": 3, "monthly": 1000000, "sheet": None},
    {"code": "PJ0018", "id": "350", "name": "Massive", "company": "株式会社MASSIVE SAPPORO", "project": "", "type": "システム開発", "amount": 14800000, "start": "2026/01/01", "end": "2026/07/30", "months": 7, "monthly": 2114286, "sheet": "Massive"},
]

# ─── PL詳細データ（個別シート用） ───
PL_DETAILS = {
    "WBT": {
        "summary": {"売上": 9475000, "原価": 1200000, "粗利": 8275000, "営利率": "87.3%"},
        "team": {"PM": "斉藤", "エンジニア": "笠井, 一谷"},
        "status": {"契約締結": False, "発注書回収": False, "請求書発行": False, "計上": False},
        "months": ["2026/1", "2026/2", "2026/3"],
        "monthly_sales": [3158333, 3158333, 3158333],
        "monthly_cost": [400000, 400000, 400000],
        "monthly_internal": [400000, 400000, 400000],
        "monthly_external": [0, 0, 0],
        "internal_members": [{"name": "斉藤", "pct": [50, 50, 50]}, {"name": "笠井", "pct": [10, 10, 10]}, {"name": "一谷", "pct": [10, 10, 10]}],
        "vendors": [],
    },
    "ASAP": {
        "summary": {"売上": 2800000, "原価": 2400000, "粗利": 400000, "営利率": "14.3%"},
        "team": {},
        "status": {"契約締結": True, "発注書回収": True, "請求書発行": True, "計上": True},
        "months": ["2025/12", "2026/1"],
        "monthly_sales": [1400000, 1400000],
        "monthly_cost": [1200000, 1200000],
        "monthly_internal": [1200000, 1200000],
        "monthly_external": [0, 0],
        "internal_members": [{"name": "（社内）", "pct": [150, 150]}],
        "vendors": [],
    },
    "Massive": {
        "summary": {"売上": 14800000, "原価": 11140000, "粗利": 3660000, "営利率": "24.7%"},
        "team": {"PM": "斉藤", "エンジニア": "佐藤, 加藤, 岩本"},
        "status": {"契約締結": False, "発注書回収": False, "請求書発行": False, "計上": False},
        "months": ["2026/2", "2026/3", "2026/4", "2026/5", "2026/6", "2026/7"],
        "monthly_sales": [2114286, 2114286, 2114286, 2114286, 2114286, 2114286],
        "monthly_cost": [550000, 2140000, 2140000, 2140000, 2140000, 2030000],
        "monthly_internal": [0, 160000, 160000, 160000, 160000, 0],
        "monthly_external": [550000, 1480000, 1480000, 1480000, 1480000, 1530000],
        "internal_members": [{"name": "佐藤", "pct": [0, 20, 20, 20, 20, 20]}],
        "vendors": [
            {"name": "斉藤", "amounts": [550000, 500000, 500000, 500000, 500000, 550000]},
            {"name": "岩本", "amounts": [0, 500000, 500000, 500000, 500000, 500000]},
            {"name": "加藤", "amounts": [0, 480000, 480000, 480000, 480000, 480000]},
        ],
    },
    "清田軌道": {
        "summary": {"売上": 4550000, "原価": 400000, "粗利": 4150000, "営利率": "91.2%"},
        "team": {},
        "status": {"契約締結": False, "発注書回収": False, "請求書発行": False, "計上": False},
        "months": ["2026/1"],
        "monthly_sales": [4550000],
        "monthly_cost": [400000],
        "monthly_internal": [400000],
        "monthly_external": [0],
        "internal_members": [{"name": "（社内）", "pct": [50]}],
        "vendors": [],
    },
    "ケイズ": {
        "summary": {"売上": 3500000, "原価": 4000000, "粗利": -500000, "営利率": "-14.3%"},
        "team": {},
        "status": {"契約締結": False, "発注書回収": False, "請求書発行": False, "計上": False},
        "months": ["2025/10", "2025/11", "2025/12", "2026/1", "2026/2", "2026/3", "2026/4"],
        "monthly_sales": [500000]*7,
        "monthly_cost": [800000]*7,
        "monthly_internal": [800000]*7,
        "monthly_external": [0]*7,
        "internal_members": [{"name": "（社内）", "pct": [100]*7}],
        "vendors": [],
    },
    "ユーグレナ": {
        "summary": {"売上": 0, "原価": 0, "粗利": 0, "営利率": "-"},
        "team": {},
        "status": {"契約締結": False, "発注書回収": False, "請求書発行": False, "計上": False},
        "months": [],
        "monthly_sales": [],
        "monthly_cost": [],
        "monthly_internal": [],
        "monthly_external": [],
        "internal_members": [],
        "vendors": [],
    },
}

# ─── 工数メンバーデータ ───
MEMBER_DATA = [
    {"no": 1, "name": "武石幸之助", "company": "", "type": "直雇用", "role": "代表", "cost_type": "原価", "mgmt": "社内"},
    {"no": 2, "name": "斎藤真介", "company": "合同会社青橋", "type": "業務委託（法人）", "role": "PM", "cost_type": "原価", "mgmt": "社内"},
    {"no": 3, "name": "橋本歩", "company": "合同会社青橋", "type": "業務委託（法人）", "role": "PMO", "cost_type": "原価", "mgmt": "社内"},
    {"no": 4, "name": "塩津瑠威", "company": "合同会社青橋", "type": "業務委託（法人）", "role": "人事サポート", "cost_type": "原価", "mgmt": "社内"},
    {"no": 5, "name": "斎藤真介", "company": "合同会社青橋", "type": "業務委託（法人）", "role": "営業/経営管理", "cost_type": "販管（管理）", "mgmt": "社内"},
    {"no": 8, "name": "笠井秀行", "company": "株式会社アイビーズ", "type": "業務委託（法人）", "role": "PM", "cost_type": "原価", "mgmt": "社内"},
    {"no": 9, "name": "加藤祐也", "company": "株式会社ワンオブゼム", "type": "業務委託（法人）", "role": "PM", "cost_type": "原価", "mgmt": "社外"},
    {"no": 10, "name": "N/A", "company": "株式会社HBLジャパン", "type": "業務委託（法人）", "role": "PM", "cost_type": "原価", "mgmt": "社外"},
    {"no": 13, "name": "森隆司", "company": "個人", "type": "業務委託（個人）", "role": "PM", "cost_type": "原価", "mgmt": "社外"},
    {"no": 14, "name": "太田ミハイル", "company": "個人", "type": "業務委託（個人）", "role": "PM", "cost_type": "原価", "mgmt": "社外"},
    {"no": 15, "name": "一谷幸一", "company": "合同会社一谷コンサルティング", "type": "業務委託（法人）", "role": "営業", "cost_type": "販管（外注）", "mgmt": "社内"},
    {"no": 16, "name": "里見吉優", "company": "株式会社TeamDoor", "type": "業務委託（法人）", "role": "営業", "cost_type": "販管（外注）", "mgmt": "社内"},
    {"no": 17, "name": "川瀬友希", "company": "株式会社Next Leap", "type": "業務委託（法人）", "role": "営業", "cost_type": "販管（外注）", "mgmt": "社内"},
    {"no": 18, "name": "木戸亜矢子", "company": "個人", "type": "業務委託（個人）", "role": "PM", "cost_type": "販管（外注）", "mgmt": "社外"},
    {"no": 19, "name": "山本裕伸", "company": "株式会社Digital Ureska", "type": "業務委託（個人）", "role": "講師", "cost_type": "販管（外注）", "mgmt": "社外"},
    {"no": 20, "name": "西村未鈴", "company": "個人", "type": "業務委託（個人）", "role": "営業", "cost_type": "販管（外注）", "mgmt": "社外"},
    {"no": 21, "name": "佐藤慶和", "company": "個人", "type": "業務委託（個人）", "role": "PM", "cost_type": "販管（外注）", "mgmt": "社内"},
    {"no": 22, "name": "森小百合", "company": "森小百合税理士事務所", "type": "業務委託（個人）", "role": "バックオフィス", "cost_type": "販管（管理）", "mgmt": "社内"},
    {"no": 23, "name": "林実咲", "company": "株式会社ワンオブゼム", "type": "OOM出向", "role": "営業、人事、広報", "cost_type": "販管（支手）", "mgmt": "社内"},
    {"no": 24, "name": "下村真由", "company": "株式会社ワンオブゼム", "type": "OOM出向", "role": "アルバイト（契約）", "cost_type": "販管（支手）", "mgmt": "社内"},
    {"no": 25, "name": "藤本健太", "company": "株式会社ワンオブゼム", "type": "OOM出向", "role": "アルバイト（営業）", "cost_type": "販管（支手）", "mgmt": "社内"},
    {"no": 26, "name": "石田愛理", "company": "株式会社ワンオブゼム", "type": "OOM出向", "role": "アルバイト（営業）", "cost_type": "販管（支手）", "mgmt": "社内"},
    {"no": 27, "name": "藤村真里菜", "company": "株式会社ワンオブゼム", "type": "OOM出向", "role": "アルバイト（レポート作成）", "cost_type": "販管（支手）", "mgmt": "社内"},
    {"no": 28, "name": "乗松美緒", "company": "合同会社queue blanche", "type": "業務委託（法人）", "role": "経営管理", "cost_type": "販管（管理）", "mgmt": "社外"},
]

# ─── 外注先データ ───
VENDOR_DATA = [
    {"name": "TF", "fixed": "", "project": "TMC", "amount": 2100000, "m1": 2100000, "m2": 2100000, "m3": 2100000},
    {"name": "株式会社nibi", "fixed": "", "project": "TMC", "amount": 500000, "m1": 500000, "m2": 500000, "m3": 0},
    {"name": "株式会社HBLジャパン", "fixed": "変動", "project": "", "amount": 0},
    {"name": "株式会社KAGEMUSHA", "fixed": "変動？", "project": "", "amount": 0},
    {"name": "株式会社ＺＹ", "fixed": "固定", "project": "", "amount": 0},
    {"name": "株式会社アイビーズ", "fixed": "固定", "project": "", "amount": 0},
    {"name": "森隆司", "fixed": "変動？", "project": "", "amount": 0},
    {"name": "太田ミハイル", "fixed": "固定", "project": "", "amount": 0},
]


def fmt_yen(v):
    """数値を円表記にする"""
    if v == 0: return "¥0"
    if v < 0: return f"-¥{abs(v):,.0f}"
    return f"¥{v:,.0f}"

def fmt_pct(v):
    """数値を%表記にする"""
    return f"{v:.1f}%"


# ═══════════════════════════════════════════════
# スプレッドシート作成
# ═══════════════════════════════════════════════

# 個別PLがあるPJ一覧
pl_sheets = [p for p in PJ_DATA if p["sheet"] and p["sheet"] in PL_DETAILS]

# シート定義
sheet_defs = [
    {"title": "ダッシュボード", "id": 0},
    {"title": "PJ台帳", "id": 1},
]
for i, p in enumerate(pl_sheets):
    sheet_defs.append({"title": f"PL_{p['name']}", "id": 100 + i})
sheet_defs.append({"title": "工数管理", "id": 50})
sheet_defs.append({"title": "外注費", "id": 51})

body = {
    "properties": {"title": "PJ PL管理（整理版）", "locale": "ja_JP"},
    "sheets": [
        {"properties": {"sheetId": s["id"], "title": s["title"], "gridProperties": {"rowCount": 200, "columnCount": 30}}}
        for s in sheet_defs
    ],
}

ss = sheets_svc.spreadsheets().create(body=body).execute()
NEW_SSID = ss["spreadsheetId"]
print(f"新規スプレッドシート作成: {NEW_SSID}")


# ═══════════════════════════════════════════════
# データ書き込み
# ═══════════════════════════════════════════════
all_data = {}

# ─── ダッシュボード ───
dash_rows = [
    ["PJ PL管理 ダッシュボード", "", "", "", "", "", "", "", "", "", "", "", ""],
    [],
    ["PJコード", "PJ名", "会社名", "商材", "契約額", "契約期間", "月額売上", "売上合計", "原価合計", "粗利", "粗利率", "ステータス", "備考"],
]
for p in PJ_DATA:
    pl = PL_DETAILS.get(p["sheet"], {})
    summary = pl.get("summary", {})
    status_dict = pl.get("status", {})

    sales = summary.get("売上", p["amount"])
    cost = summary.get("原価", 0)
    profit = summary.get("粗利", 0)
    rate = summary.get("営利率", "-")

    # ステータス判定
    if status_dict:
        done = sum(1 for v in status_dict.values() if v)
        if done == 4: status = "完了"
        elif done > 0: status = "進行中"
        else: status = "未着手"
    else:
        status = ""

    period = ""
    if p["start"] and p["end"]:
        period = f"{p['start']}～{p['end']}"

    dash_rows.append([
        p["code"], p["name"], p["company"], p["type"],
        fmt_yen(p["amount"]) if p["amount"] else "",
        period,
        fmt_yen(p["monthly"]) if p["monthly"] else "",
        fmt_yen(sales) if sales else "",
        fmt_yen(cost) if cost else "",
        fmt_yen(profit) if profit else "",
        rate if rate else "",
        status,
        "",
    ])

all_data["ダッシュボード!A1"] = dash_rows

# ─── PJ台帳 ───
ledger_rows = [
    ["PJ台帳"],
    [],
    ["PJコード", "ID", "PJ名", "会社名", "プロジェクト名", "商材", "契約額", "契約開始日", "契約終了日", "契約期間", "月額売上", "メモ"],
]
for p in PJ_DATA:
    ledger_rows.append([
        p["code"], p["id"], p["name"], p["company"], p["project"], p["type"],
        fmt_yen(p["amount"]) if p["amount"] else "",
        p["start"], p["end"],
        f"{p['months']}ヶ月" if p["months"] else "",
        fmt_yen(p["monthly"]) if p["monthly"] else "",
        "",
    ])

all_data["PJ台帳!A1"] = ledger_rows

# ─── 個別PLシート ───
for p in pl_sheets:
    pl = PL_DETAILS[p["sheet"]]
    s = pl["summary"]
    sheet_name = f"PL_{p['name']}"
    months = pl["months"]
    n_months = len(months)

    rows = []
    # 基本情報
    rows.append([f"PL: {p['name']}", "", "", "会社名:", p["company"]])
    rows.append(["商材:", p["type"], "", "契約額:", fmt_yen(p["amount"])])
    period = f"{p['start']}～{p['end']}" if p["start"] else ""
    rows.append(["契約期間:", period, "", "月額売上:", fmt_yen(p["monthly"])])

    # チーム
    team_str = "  ".join(f"{k}: {v}" for k, v in pl.get("team", {}).items())
    rows.append(["PJチーム:", team_str])

    # ステータス
    st = pl.get("status", {})
    status_items = [f"{k}: {'✓' if v else '—'}" for k, v in st.items()]
    rows.append(["ステータス:", "  ".join(status_items)])

    rows.append([])  # 空行

    # サマリー
    profit = s["売上"] - s["原価"]
    rate_val = (profit / s["売上"] * 100) if s["売上"] else 0
    rows.append(["【サマリー】", "売上高", "原価", "粗利", "粗利率", "営利"])
    rows.append(["", fmt_yen(s["売上"]), fmt_yen(s["原価"]), fmt_yen(s.get("粗利", profit)), s["営利率"], fmt_yen(s.get("粗利", profit))])

    rows.append([])  # 空行

    # 月別計上
    if months:
        header = ["【月別計上】", "科目"] + months + ["合計"]
        rows.append(header)

        def sum_row(vals):
            return sum(v for v in vals if isinstance(v, (int, float)))

        ms = pl["monthly_sales"]
        mc = pl["monthly_cost"]
        mi = pl["monthly_internal"]
        me = pl["monthly_external"]

        rows.append(["", "売上高"] + [fmt_yen(v) for v in ms] + [fmt_yen(sum_row(ms))])
        rows.append(["", "原価"] + [fmt_yen(v) for v in mc] + [fmt_yen(sum_row(mc))])
        rows.append(["", "  社内工数"] + [fmt_yen(v) for v in mi] + [fmt_yen(sum_row(mi))])
        rows.append(["", "  外注費"] + [fmt_yen(v) for v in me] + [fmt_yen(sum_row(me))])
        gross = [ms[i] - mc[i] for i in range(n_months)]
        rows.append(["", "粗利"] + [fmt_yen(v) for v in gross] + [fmt_yen(sum_row(gross))])
        gross_pct = [fmt_pct(gross[i]/ms[i]*100) if ms[i] else "-" for i in range(n_months)]
        rows.append(["", "粗利率"] + gross_pct + [s["営利率"]])

        rows.append([])

        # 社内工数内訳
        if pl["internal_members"]:
            rows.append(["【社内工数内訳】", "メンバー", "単価: ¥800,000"] + months)
            for m in pl["internal_members"]:
                pcts = [f"{v}%" for v in m["pct"]]
                rows.append(["", m["name"]] + [""] + pcts)

        rows.append([])

        # 外注費内訳
        if pl["vendors"]:
            rows.append(["【外注費内訳】", "ベンダー"] + months + ["合計"])
            for v in pl["vendors"]:
                rows.append(["", v["name"]] + [fmt_yen(a) for a in v["amounts"]] + [fmt_yen(sum_row(v["amounts"]))])
            total_vendor = [sum(v["amounts"][i] for v in pl["vendors"]) for i in range(n_months)]
            rows.append(["", "小計"] + [fmt_yen(v) for v in total_vendor] + [fmt_yen(sum_row(total_vendor))])

    all_data[f"{sheet_name}!A1"] = rows

# ─── 工数管理 ───
member_rows = [
    ["工数管理", "", "", "", "", "", "", "単価: ¥800,000"],
    [],
    ["No", "氏名", "会社名", "雇用形態", "役割", "原価/販管", "管理形態", "メモ"],
]
for m in MEMBER_DATA:
    member_rows.append([
        m["no"], m["name"], m["company"], m["type"], m["role"], m["cost_type"], m["mgmt"], "",
    ])

all_data["工数管理!A1"] = member_rows

# ─── 外注費 ───
vendor_rows = [
    ["外注費管理"],
    [],
    ["契約先", "固定/変動", "プロジェクト", "月額", "1月", "2月", "3月", "4月", "5月", "6月", "合計", "メモ"],
]
for v in VENDOR_DATA:
    row = [v["name"], v["fixed"], v.get("project", ""), fmt_yen(v["amount"]) if v["amount"] else ""]
    if "m1" in v:
        row += [fmt_yen(v.get("m1", 0)), fmt_yen(v.get("m2", 0)), fmt_yen(v.get("m3", 0))]
    vendor_rows.append(row)

all_data["外注費!A1"] = vendor_rows


# 一括書き込み
batch_data = []
for rng, vals in all_data.items():
    batch_data.append({"range": rng, "values": vals})

sheets_svc.spreadsheets().values().batchUpdate(
    spreadsheetId=NEW_SSID,
    body={"valueInputOption": "RAW", "data": batch_data},
).execute()
print("データ書き込み完了")


# ═══════════════════════════════════════════════
# 書式設定
# ═══════════════════════════════════════════════
fmt_requests = []

def cell_fmt(sheet_id, r1, r2, c1, c2, bg=None, bold=False, fg=None, h_align=None, font_size=None):
    """セル書式設定リクエストを作成"""
    fmt = {}
    if bg: fmt["backgroundColor"] = bg
    text_fmt = {}
    if bold: text_fmt["bold"] = True
    if fg: text_fmt["foregroundColor"] = fg
    if font_size: text_fmt["fontSize"] = font_size
    if text_fmt: fmt["textFormat"] = text_fmt
    if h_align: fmt["horizontalAlignment"] = h_align
    fields = []
    if "backgroundColor" in fmt: fields.append("userEnteredFormat.backgroundColor")
    if "textFormat" in fmt: fields.append("userEnteredFormat.textFormat")
    if "horizontalAlignment" in fmt: fields.append("userEnteredFormat.horizontalAlignment")
    return {
        "repeatCell": {
            "range": {"sheetId": sheet_id, "startRowIndex": r1, "endRowIndex": r2, "startColumnIndex": c1, "endColumnIndex": c2},
            "cell": {"userEnteredFormat": fmt},
            "fields": ",".join(fields),
        }
    }

def col_width(sheet_id, col, width):
    return {
        "updateDimensionProperties": {
            "range": {"sheetId": sheet_id, "dimension": "COLUMNS", "startIndex": col, "endIndex": col + 1},
            "properties": {"pixelSize": width},
            "fields": "pixelSize",
        }
    }

def freeze(sheet_id, rows=0, cols=0):
    props = {}
    if rows: props["frozenRowCount"] = rows
    if cols: props["frozenColumnCount"] = cols
    return {
        "updateSheetProperties": {
            "properties": {"sheetId": sheet_id, "gridProperties": props},
            "fields": ",".join(f"gridProperties.{k}" for k in props),
        }
    }

# ─── ダッシュボード書式 ───
sid = 0
# タイトル
fmt_requests.append(cell_fmt(sid, 0, 1, 0, 13, bg=DARK_BLUE, bold=True, fg=WHITE, font_size=14))
# ヘッダー
fmt_requests.append(cell_fmt(sid, 2, 3, 0, 13, bg=MID_BLUE, bold=True, fg=WHITE))
# データ行（交互背景）
for i in range(len(PJ_DATA)):
    row = 3 + i
    bg = LIGHT_GRAY if i % 2 == 0 else WHITE
    fmt_requests.append(cell_fmt(sid, row, row+1, 0, 13, bg=bg))
# 列幅
for c, w in [(0, 90), (1, 160), (2, 250), (3, 100), (4, 120), (5, 200), (6, 110), (7, 110), (8, 110), (9, 110), (10, 80), (11, 80), (12, 120)]:
    fmt_requests.append(col_width(sid, c, w))
fmt_requests.append(freeze(sid, rows=3))

# ─── PJ台帳書式 ───
sid = 1
fmt_requests.append(cell_fmt(sid, 0, 1, 0, 12, bg=DARK_BLUE, bold=True, fg=WHITE, font_size=14))
fmt_requests.append(cell_fmt(sid, 2, 3, 0, 12, bg=MID_BLUE, bold=True, fg=WHITE))
for i in range(len(PJ_DATA)):
    row = 3 + i
    bg = LIGHT_GRAY if i % 2 == 0 else WHITE
    fmt_requests.append(cell_fmt(sid, row, row+1, 0, 12, bg=bg))
for c, w in [(0, 90), (1, 50), (2, 160), (3, 250), (4, 250), (5, 100), (6, 120), (7, 110), (8, 110), (9, 80), (10, 110), (11, 150)]:
    fmt_requests.append(col_width(sid, c, w))
fmt_requests.append(freeze(sid, rows=3))

# ─── 個別PLシート書式 ───
for i, p in enumerate(pl_sheets):
    sid = 100 + i
    # ヘッダー部（行1-5）
    fmt_requests.append(cell_fmt(sid, 0, 1, 0, 15, bg=DARK_BLUE, bold=True, fg=WHITE, font_size=12))
    fmt_requests.append(cell_fmt(sid, 1, 5, 0, 1, bold=True))
    # サマリーヘッダー
    fmt_requests.append(cell_fmt(sid, 6, 7, 0, 6, bg=LIGHT_GREEN, bold=True))
    # 月別計上ヘッダー
    fmt_requests.append(cell_fmt(sid, 9, 10, 0, 15, bg=LIGHT_BLUE, bold=True))
    # 列幅
    fmt_requests.append(col_width(sid, 0, 130))
    fmt_requests.append(col_width(sid, 1, 110))
    for c in range(2, 15):
        fmt_requests.append(col_width(sid, c, 100))

# ─── 工数管理書式 ───
sid = 50
fmt_requests.append(cell_fmt(sid, 0, 1, 0, 8, bg=DARK_BLUE, bold=True, fg=WHITE, font_size=14))
fmt_requests.append(cell_fmt(sid, 2, 3, 0, 8, bg=MID_BLUE, bold=True, fg=WHITE))
for i in range(len(MEMBER_DATA)):
    row = 3 + i
    bg = LIGHT_GRAY if i % 2 == 0 else WHITE
    fmt_requests.append(cell_fmt(sid, row, row+1, 0, 8, bg=bg))
for c, w in [(0, 40), (1, 130), (2, 220), (3, 150), (4, 160), (5, 110), (6, 80), (7, 150)]:
    fmt_requests.append(col_width(sid, c, w))
fmt_requests.append(freeze(sid, rows=3))

# ─── 外注費書式 ───
sid = 51
fmt_requests.append(cell_fmt(sid, 0, 1, 0, 12, bg=DARK_BLUE, bold=True, fg=WHITE, font_size=14))
fmt_requests.append(cell_fmt(sid, 2, 3, 0, 12, bg=MID_BLUE, bold=True, fg=WHITE))
for c, w in [(0, 180), (1, 80), (2, 100), (3, 100)]:
    fmt_requests.append(col_width(sid, c, w))
fmt_requests.append(freeze(sid, rows=3))

# 一括適用
sheets_svc.spreadsheets().batchUpdate(
    spreadsheetId=NEW_SSID,
    body={"requests": fmt_requests},
).execute()
print("書式設定完了")

# 共有設定（リンクを知っている人が閲覧可能）
drive_svc.permissions().create(
    fileId=NEW_SSID,
    body={"role": "writer", "type": "anyone"},
).execute()
print("共有設定完了")

print(f"\n✅ 完成！")
print(f"https://docs.google.com/spreadsheets/d/{NEW_SSID}")
