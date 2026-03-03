#!/usr/bin/env python3
"""
PJ PL管理スプレッドシート v2 — 数式＋クロスシート連携版
"""
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

TOKEN = "/Users/shinsukesaito/claude/ai-cost-monitor/token.json"
creds = Credentials.from_authorized_user_file(TOKEN)
sheets_svc = build("sheets", "v4", credentials=creds)
drive_svc = build("drive", "v3", credentials=creds)

# ─── 色 ───
WHITE = {"red": 1, "green": 1, "blue": 1}
DARK_BLUE = {"red": 0.15, "green": 0.25, "blue": 0.45}
LIGHT_BLUE = {"red": 0.85, "green": 0.92, "blue": 1.0}
LIGHT_GRAY = {"red": 0.95, "green": 0.95, "blue": 0.95}
LIGHT_GREEN = {"red": 0.85, "green": 0.95, "blue": 0.85}
LIGHT_YELLOW = {"red": 1.0, "green": 0.98, "blue": 0.85}
MID_BLUE = {"red": 0.6, "green": 0.75, "blue": 0.95}
LIGHT_RED = {"red": 1.0, "green": 0.92, "blue": 0.92}

# ─── PJデータ ───
# sheet_key: PL_SHEETSのキーに対応
PJ_LIST = [
    {"code": "PJ0001", "name": "TMC（WBT）", "company": "ウーブン・バイ・トヨタ株式会社", "type": "コンサル", "amount": 9475000, "start": "2026/01/05", "end": "2026/03/13", "months": 3, "monthly": 3158333, "key": "WBT"},
    {"code": "PJ0002", "name": "ASAP", "company": "ASAP SECURITY 株式会社", "type": "システム開発", "amount": 2800000, "start": "2025/11/01", "end": "2026/02/28", "months": 4, "monthly": 700000, "key": "ASAP"},
    {"code": "PJ0005", "name": "清田軌道", "company": "清田軌道工業（株）", "type": "研修", "amount": 4550000, "start": "2026/01/27", "end": "2026/01/28", "months": 1, "monthly": 4550000, "key": "清田軌道"},
    {"code": "PJ0006", "name": "ケイズ", "company": "株式会社ケイズグループ", "type": "その他", "amount": 3500000, "start": "2025/10/01", "end": "2026/04/30", "months": 7, "monthly": 500000, "key": "ケイズ"},
    {"code": "PJ0018", "name": "Massive", "company": "株式会社MASSIVE SAPPORO", "type": "システム開発", "amount": 14800000, "start": "2026/01/01", "end": "2026/07/30", "months": 7, "monthly": 2114286, "key": "Massive"},
    {"code": "PJ0004", "name": "ユーグレナ", "company": "株式会社ユーグレナ", "type": "提携", "amount": 0, "start": "", "end": "", "months": 1, "monthly": 0, "key": "ユーグレナ"},
    {"code": "PJ0007", "name": "栗田工業（IST）", "company": "栗田工業株式会社", "type": "研修", "amount": 3900000, "start": "2025/11/01", "end": "2026/02/28", "months": 4, "monthly": 975000, "key": "栗田工業"},
    {"code": "PJ0008", "name": "小野薬品（AIMG）", "company": "小野薬品工業 株式会社", "type": "研修", "amount": 1200000, "start": "2026/02/01", "end": "2026/02/28", "months": 1, "monthly": 1200000, "key": "小野薬品"},
    {"code": "PJ0011", "name": "パナソニック", "company": "パナソニック株式会社", "type": "研修", "amount": 3000000, "start": "2026/02/01", "end": "2026/03/31", "months": 2, "monthly": 1500000, "key": "パナソニック"},
    {"code": "PJ0017", "name": "ONE COMPATH", "company": "株式会社ONE COMPATH", "type": "システム開発", "amount": 3000000, "start": "2026/01/01", "end": "2026/03/31", "months": 3, "monthly": 1000000, "key": "ONE_COMPATH"},
]

# PL詳細（月別の初期値）— 数値のまま保持
PL_SHEETS = {
    "WBT": {
        "month_labels": ["2026/1", "2026/2", "2026/3"],
        "sales":     [3158333, 3158333, 3158333],
        "members":   [
            {"name": "斉藤", "pct": [50, 50, 50]},
            {"name": "笠井", "pct": [10, 10, 10]},
            {"name": "一谷", "pct": [10, 10, 10]},
        ],
        "vendors": [],
        "status": {"締結": False, "発注書": False, "請求書": False, "計上": False},
        "team": "PM: 斉藤  エンジニア: 笠井, 一谷",
    },
    "ASAP": {
        "month_labels": ["2025/12", "2026/1"],
        "sales":     [1400000, 1400000],
        "members":   [{"name": "（社内）", "pct": [150, 150]}],
        "vendors": [],
        "status": {"締結": True, "発注書": True, "請求書": True, "計上": True},
        "team": "",
    },
    "清田軌道": {
        "month_labels": ["2026/1"],
        "sales":     [4550000],
        "members":   [{"name": "（社内）", "pct": [50]}],
        "vendors": [],
        "status": {"締結": False, "発注書": False, "請求書": False, "計上": False},
        "team": "",
    },
    "ケイズ": {
        "month_labels": ["2025/10", "2025/11", "2025/12", "2026/1", "2026/2", "2026/3", "2026/4"],
        "sales":     [500000]*7,
        "members":   [{"name": "（社内）", "pct": [100]*7}],
        "vendors": [],
        "status": {"締結": False, "発注書": False, "請求書": False, "計上": False},
        "team": "",
    },
    "Massive": {
        "month_labels": ["2026/2", "2026/3", "2026/4", "2026/5", "2026/6", "2026/7"],
        "sales":     [2114286, 2114286, 2114286, 2114286, 2114286, 2114286],
        "members":   [{"name": "佐藤", "pct": [0, 20, 20, 20, 20, 20]}],
        "vendors": [
            {"name": "斉藤", "amounts": [550000, 500000, 500000, 500000, 500000, 550000]},
            {"name": "岩本", "amounts": [0, 500000, 500000, 500000, 500000, 500000]},
            {"name": "加藤", "amounts": [0, 480000, 480000, 480000, 480000, 480000]},
        ],
        "status": {"締結": False, "発注書": False, "請求書": False, "計上": False},
        "team": "PM: 斉藤  エンジニア: 佐藤, 加藤, 岩本",
    },
    "ユーグレナ": {
        "month_labels": ["2026/1", "2026/2", "2026/3"],
        "sales":     [0, 0, 0],
        "members":   [],
        "vendors": [],
        "status": {"締結": False, "発注書": False, "請求書": False, "計上": False},
        "team": "",
    },
}

# PLがない案件もシートだけ作る
for pj in PJ_LIST:
    if pj["key"] not in PL_SHEETS:
        n = max(pj["months"], 3)
        PL_SHEETS[pj["key"]] = {
            "month_labels": [f"月{i+1}" for i in range(n)],
            "sales": [pj["monthly"]] * n if pj["monthly"] else [0] * n,
            "members": [],
            "vendors": [],
            "status": {"締結": False, "発注書": False, "請求書": False, "計上": False},
            "team": "",
        }

# ─── 固定レイアウト定数 ───
# 個別PLシートの行番号（0-indexed for API, 1-indexed for formulas）
R_TITLE = 0      # タイトル
R_INFO = 1       # 基本情報（3行）
R_EMPTY1 = 4
R_SUMMARY_H = 5  # サマリーヘッダー
R_SUMMARY_V = 6  # サマリー値
R_EMPTY2 = 7
R_MONTHLY_H = 8  # 月別ヘッダー
R_SALES = 9      # 売上
R_COST = 10      # 原価（数式）
R_INTERNAL = 11  # 社内工数（数式）
R_EXTERNAL = 12  # 外注費（数式）
R_GROSS = 13     # 粗利（数式）
R_GROSS_PCT = 14 # 粗利率（数式）
R_EMPTY3 = 15
R_LABOR_H = 16   # 工数ヘッダー
R_LABOR_START = 17  # メンバー開始（最大5名）
MAX_MEMBERS = 5
R_LABOR_SUB = R_LABOR_START + MAX_MEMBERS  # 工数小計 = 22
R_EMPTY4 = R_LABOR_SUB + 1  # 23
R_VENDOR_H = R_EMPTY4 + 1   # 外注ヘッダー = 24
R_VENDOR_START = R_VENDOR_H + 1  # ベンダー開始 = 25（最大5社）
MAX_VENDORS = 5
R_VENDOR_SUB = R_VENDOR_START + MAX_VENDORS  # 外注小計 = 30

UNIT_PRICE = 800000  # 社内工数単価
L_START = 3  # 工数セクションの月データ開始列（D列）— C列は単価のため1列ずれる

def col_letter(idx):
    """0-indexed列番号をアルファベットに変換"""
    if idx < 26: return chr(65 + idx)
    return chr(64 + idx // 26) + chr(65 + idx % 26)

def R(row_0indexed):
    """0-indexed → 1-indexed（数式用）"""
    return row_0indexed + 1


# ═══════════════════════════════════════════════
# スプレッドシート作成
# ═══════════════════════════════════════════════
sheet_defs = [{"title": "ダッシュボード", "id": 0}]
pl_sheet_map = {}  # key → (sheet_title, sheet_id)
for i, pj in enumerate(PJ_LIST):
    title = f"PL_{pj['name']}"
    sid = 100 + i
    sheet_defs.append({"title": title, "id": sid})
    pl_sheet_map[pj["key"]] = (title, sid)

sheet_defs.append({"title": "工数管理", "id": 50})
sheet_defs.append({"title": "外注費管理", "id": 51})
sheet_defs.append({"title": "PJ台帳", "id": 1})

body = {
    "properties": {"title": "PJ PL管理 v2（数式版）", "locale": "ja_JP"},
    "sheets": [
        {"properties": {"sheetId": s["id"], "title": s["title"], "gridProperties": {"rowCount": 100, "columnCount": 30}}}
        for s in sheet_defs
    ],
}
ss = sheets_svc.spreadsheets().create(body=body).execute()
NEW_SSID = ss["spreadsheetId"]
print(f"作成: {NEW_SSID}")


# ═══════════════════════════════════════════════
# 個別PLシート — データ＋数式を書き込み
# ═══════════════════════════════════════════════
all_values = {}

for pj in PJ_LIST:
    key = pj["key"]
    pl = PL_SHEETS[key]
    title = pl_sheet_map[key][0]
    months = pl["month_labels"]
    N = len(months)
    # 月の列: C(2)〜C+N-1, 合計列: C+N
    C_START = 2  # 月データ開始列（C列）
    C_END = C_START + N  # 合計列
    c_sum = col_letter(C_END)

    rows = [None] * (R_VENDOR_SUB + 2)

    # Row 0: タイトル
    rows[R_TITLE] = [f"PL: {pj['name']}", "", ""] + [""] * N

    # Row 1-3: 基本情報
    rows[R_INFO] = ["会社名", pj["company"], "", "商材", pj["type"]]
    rows[R_INFO + 1] = ["契約額", pj["amount"], "", "月額売上", pj["monthly"]]
    period = f"{pj['start']}～{pj['end']}" if pj["start"] else ""
    rows[R_INFO + 2] = ["契約期間", period, "", "チーム", pl.get("team", "")]

    # Row 4: empty
    rows[R_EMPTY1] = []

    # Row 5: サマリーヘッダー
    rows[R_SUMMARY_H] = ["サマリー", "売上高", "原価", "粗利", "粗利率", "営利"]

    # Row 6: サマリー値（数式）— 月別の合計を参照
    r_s = R(R_SALES)  # 売上行（1-indexed）
    r_c = R(R_COST)
    r_g = R(R_GROSS)
    c0 = col_letter(C_START)
    cn = col_letter(C_END - 1)
    rows[R_SUMMARY_V] = [
        "",
        f"=SUM({c0}{r_s}:{cn}{r_s})",   # 売上合計
        f"=SUM({c0}{r_c}:{cn}{r_c})",   # 原価合計
        f"=B{R(R_SUMMARY_V)}-C{R(R_SUMMARY_V)}",  # 粗利
        f'=IF(B{R(R_SUMMARY_V)}=0,"-",D{R(R_SUMMARY_V)}/B{R(R_SUMMARY_V)})',  # 粗利率
        f"=D{R(R_SUMMARY_V)}",  # 営利＝粗利（販管費なしの場合）
    ]

    # Row 7: empty
    rows[R_EMPTY2] = []

    # Row 8: 月別ヘッダー
    rows[R_MONTHLY_H] = ["月別計上", "科目"] + months + ["合計"]

    # Row 9: 売上（初期値を入力）
    sales_row = ["", "売上高"]
    for v in pl["sales"]:
        sales_row.append(v)
    sales_row.append(f"=SUM({c0}{R(R_SALES)}:{cn}{R(R_SALES)})")
    rows[R_SALES] = sales_row

    # Row 10: 原価（= 社内工数 + 外注費）
    cost_row = ["", "原価"]
    for i in range(N):
        c = col_letter(C_START + i)
        cost_row.append(f"={c}{R(R_INTERNAL)}+{c}{R(R_EXTERNAL)}")
    cost_row.append(f"=SUM({c0}{R(R_COST)}:{cn}{R(R_COST)})")
    rows[R_COST] = cost_row

    # Row 11: 社内工数（= 工数セクション小計を参照）
    int_row = ["", "  社内工数"]
    for i in range(N):
        c = col_letter(L_START + i)  # 工数小計のD列以降を参照
        int_row.append(f"={c}{R(R_LABOR_SUB)}")
    int_row.append(f"=SUM({c0}{R(R_INTERNAL)}:{cn}{R(R_INTERNAL)})")
    rows[R_INTERNAL] = int_row

    # Row 12: 外注費（= 外注セクション小計を参照）
    ext_row = ["", "  外注費"]
    for i in range(N):
        c = col_letter(C_START + i)
        ext_row.append(f"={c}{R(R_VENDOR_SUB)}")
    ext_row.append(f"=SUM({c0}{R(R_EXTERNAL)}:{cn}{R(R_EXTERNAL)})")
    rows[R_EXTERNAL] = ext_row

    # Row 13: 粗利（= 売上 - 原価）
    gross_row = ["", "粗利"]
    for i in range(N):
        c = col_letter(C_START + i)
        gross_row.append(f"={c}{R(R_SALES)}-{c}{R(R_COST)}")
    gross_row.append(f"=SUM({c0}{R(R_GROSS)}:{cn}{R(R_GROSS)})")
    rows[R_GROSS] = gross_row

    # Row 14: 粗利率
    pct_row = ["", "粗利率"]
    for i in range(N):
        c = col_letter(C_START + i)
        pct_row.append(f'=IF({c}{R(R_SALES)}=0,"-",{c}{R(R_GROSS)}/{c}{R(R_SALES)})')
    pct_row.append(f'=IF({c_sum}{R(R_SALES)}=0,"-",{c_sum}{R(R_GROSS)}/{c_sum}{R(R_SALES)})')
    rows[R_GROSS_PCT] = pct_row

    # Row 15: empty
    rows[R_EMPTY3] = []

    # Row 16: 工数ヘッダー
    rows[R_LABOR_H] = ["社内工数", "メンバー", "単価"] + months + ["合計(円)"]

    # Row 17-21: メンバー（最大5名、入力欄：配分率%）
    members = pl.get("members", [])
    for m_idx in range(MAX_MEMBERS):
        r = R_LABOR_START + m_idx
        if m_idx < len(members):
            m = members[m_idx]
            member_row = ["", m["name"], UNIT_PRICE]
            for i in range(N):
                member_row.append(m["pct"][i] / 100 if i < len(m["pct"]) else 0)
            # 合計(円) = SUM(単価 × 各月配分率)
            sum_parts = []
            for i in range(N):
                c = col_letter(L_START + i)  # 月データはD列から
                sum_parts.append(f"C{R(r)}*{c}{R(r)}")
            member_row.append(f"={'+'.join(sum_parts)}")
        else:
            member_row = ["", "", UNIT_PRICE] + [0] * N + [f"=0"]
        rows[r] = member_row

    # Row 22: 工数小計（= 各月のSUM(単価 × 配分率)）
    sub_row = ["", "小計", ""]
    for i in range(N):
        c = col_letter(L_START + i)  # 月データはD列から
        parts = []
        for m_idx in range(MAX_MEMBERS):
            mr = R(R_LABOR_START + m_idx)
            parts.append(f"$C{mr}*{c}{mr}")
        sub_row.append(f"={'+'.join(parts)}")
    l_c0 = col_letter(L_START)
    l_cn = col_letter(L_START + N - 1)
    sub_row.append(f"=SUM({l_c0}{R(R_LABOR_SUB)}:{l_cn}{R(R_LABOR_SUB)})")
    rows[R_LABOR_SUB] = sub_row

    # Row 23: empty
    rows[R_EMPTY4] = []

    # Row 24: 外注ヘッダー
    rows[R_VENDOR_H] = ["外注費", "ベンダー"] + months + ["合計"]

    # Row 25-29: ベンダー（最大5社、入力欄：金額）
    vendors = pl.get("vendors", [])
    for v_idx in range(MAX_VENDORS):
        r = R_VENDOR_START + v_idx
        if v_idx < len(vendors):
            v = vendors[v_idx]
            vendor_row = ["", v["name"]]
            for i in range(N):
                vendor_row.append(v["amounts"][i] if i < len(v["amounts"]) else 0)
            vendor_row.append(f"=SUM({c0}{R(r)}:{cn}{R(r)})")
        else:
            vendor_row = ["", ""] + [0] * N + [f"=SUM({c0}{R(r)}:{cn}{R(r)})"]
        rows[r] = vendor_row

    # Row 30: 外注小計
    vsub_row = ["", "小計"]
    for i in range(N):
        c = col_letter(C_START + i)
        vsub_row.append(f"=SUM({c}{R(R_VENDOR_START)}:{c}{R(R_VENDOR_START + MAX_VENDORS - 1)})")
    vsub_row.append(f"=SUM({c0}{R(R_VENDOR_SUB)}:{cn}{R(R_VENDOR_SUB)})")
    rows[R_VENDOR_SUB] = vsub_row

    # Noneを空リストに
    rows = [r if r is not None else [] for r in rows]

    all_values[f"'{title}'!A1"] = rows


# ═══════════════════════════════════════════════
# ダッシュボード — 全PLから自動集約
# ═══════════════════════════════════════════════
dash_rows = [
    ["PJ PL管理 ダッシュボード"],
    [],
    ["PJコード", "PJ名", "会社名", "商材", "契約額", "月額売上", "売上合計", "原価合計", "粗利", "粗利率", "ステータス"],
]

for pj in PJ_LIST:
    title = pl_sheet_map[pj["key"]][0]
    # 各PLシートのサマリー行から数式で参照
    dash_rows.append([
        pj["code"],
        pj["name"],
        pj["company"],
        pj["type"],
        pj["amount"],
        pj["monthly"],
        f"='{title}'!B{R(R_SUMMARY_V)}",    # 売上合計
        f"='{title}'!C{R(R_SUMMARY_V)}",    # 原価合計
        f"='{title}'!D{R(R_SUMMARY_V)}",    # 粗利
        f"='{title}'!E{R(R_SUMMARY_V)}",    # 粗利率
        "",
    ])

# 合計行
n_pj = len(PJ_LIST)
start_r = 4
end_r = start_r + n_pj - 1
dash_rows.append([])
dash_rows.append([
    "", "", "", "", f"=SUM(E{start_r}:E{end_r})", f"=SUM(F{start_r}:F{end_r})",
    f"=SUM(G{start_r}:G{end_r})", f"=SUM(H{start_r}:H{end_r})",
    f"=SUM(I{start_r}:I{end_r})",
    f'=IF(G{end_r+2}=0,"-",I{end_r+2}/G{end_r+2})',
    "",
])
dash_rows[-1][0] = "合計"

all_values["ダッシュボード!A1"] = dash_rows


# ═══════════════════════════════════════════════
# 工数管理 — 各PLから自動集約
# ═══════════════════════════════════════════════
labor_rows = [
    ["工数管理（自動集約）", "", "単価: ¥800,000"],
    [],
    ["PJ名", "メンバー", "合計工数(円)"],
]

for pj in PJ_LIST:
    title = pl_sheet_map[pj["key"]][0]
    for m_idx in range(MAX_MEMBERS):
        r = R(R_LABOR_START + m_idx)
        # メンバー名と合計金額を各PLから参照
        labor_rows.append([
            pj["name"],
            f"='{title}'!B{r}",
            f"='{title}'!{col_letter(L_START + len(PL_SHEETS[pj['key']]['month_labels']))}{r}",
        ])
    # 小計行
    sub_r = R(R_LABOR_SUB)
    cn_labor = col_letter(L_START + len(PL_SHEETS[pj['key']]['month_labels']))
    labor_rows.append([
        f"  {pj['name']} 小計", "",
        f"='{title}'!{cn_labor}{sub_r}",
    ])
    labor_rows.append([])

all_values["工数管理!A1"] = labor_rows


# ═══════════════════════════════════════════════
# 外注費管理 — 各PLから自動集約
# ═══════════════════════════════════════════════
vendor_mgmt_rows = [
    ["外注費管理（自動集約）"],
    [],
    ["PJ名", "ベンダー", "合計"],
]

for pj in PJ_LIST:
    title = pl_sheet_map[pj["key"]][0]
    for v_idx in range(MAX_VENDORS):
        r = R(R_VENDOR_START + v_idx)
        cn_vendor = col_letter(C_START + len(PL_SHEETS[pj['key']]['month_labels']))
        vendor_mgmt_rows.append([
            pj["name"],
            f"='{title}'!B{r}",
            f"='{title}'!{cn_vendor}{r}",
        ])
    sub_r = R(R_VENDOR_SUB)
    cn_vendor = col_letter(C_START + len(PL_SHEETS[pj['key']]['month_labels']))
    vendor_mgmt_rows.append([
        f"  {pj['name']} 小計", "",
        f"='{title}'!{cn_vendor}{sub_r}",
    ])
    vendor_mgmt_rows.append([])

all_values["外注費管理!A1"] = vendor_mgmt_rows


# ═══════════════════════════════════════════════
# PJ台帳
# ═══════════════════════════════════════════════
ledger_rows = [
    ["PJ台帳"],
    [],
    ["PJコード", "PJ名", "会社名", "商材", "契約額", "契約開始", "契約終了", "期間", "月額売上"],
]
for pj in PJ_LIST:
    ledger_rows.append([
        pj["code"], pj["name"], pj["company"], pj["type"],
        pj["amount"], pj["start"], pj["end"],
        f"{pj['months']}ヶ月" if pj["months"] else "", pj["monthly"],
    ])

all_values["PJ台帳!A1"] = ledger_rows


# ═══════════════════════════════════════════════
# 一括書き込み（USER_ENTERED → 数式が有効になる）
# ═══════════════════════════════════════════════
batch = [{"range": rng, "values": vals} for rng, vals in all_values.items()]
sheets_svc.spreadsheets().values().batchUpdate(
    spreadsheetId=NEW_SSID,
    body={"valueInputOption": "USER_ENTERED", "data": batch},
).execute()
print("データ＋数式書き込み完了")


# ═══════════════════════════════════════════════
# 書式設定
# ═══════════════════════════════════════════════
requests = []

def cell_fmt(sid, r1, r2, c1, c2, bg=None, bold=False, fg=None, font_size=None, num_fmt=None, h_align=None):
    fmt = {}
    fields = []
    if bg:
        fmt["backgroundColor"] = bg
        fields.append("userEnteredFormat.backgroundColor")
    text_fmt = {}
    if bold: text_fmt["bold"] = True
    if fg: text_fmt["foregroundColor"] = fg
    if font_size: text_fmt["fontSize"] = font_size
    if text_fmt:
        fmt["textFormat"] = text_fmt
        fields.append("userEnteredFormat.textFormat")
    if num_fmt:
        fmt["numberFormat"] = num_fmt
        fields.append("userEnteredFormat.numberFormat")
    if h_align:
        fmt["horizontalAlignment"] = h_align
        fields.append("userEnteredFormat.horizontalAlignment")
    return {"repeatCell": {
        "range": {"sheetId": sid, "startRowIndex": r1, "endRowIndex": r2, "startColumnIndex": c1, "endColumnIndex": c2},
        "cell": {"userEnteredFormat": fmt}, "fields": ",".join(fields),
    }}

def col_w(sid, c, w):
    return {"updateDimensionProperties": {
        "range": {"sheetId": sid, "dimension": "COLUMNS", "startIndex": c, "endIndex": c+1},
        "properties": {"pixelSize": w}, "fields": "pixelSize",
    }}

def freeze(sid, rows=0, cols=0):
    props = {}
    if rows: props["frozenRowCount"] = rows
    if cols: props["frozenColumnCount"] = cols
    return {"updateSheetProperties": {
        "properties": {"sheetId": sid, "gridProperties": props},
        "fields": ",".join(f"gridProperties.{k}" for k in props),
    }}

YEN_FMT = {"type": "NUMBER", "pattern": "¥#,##0"}
PCT_FMT = {"type": "PERCENT", "pattern": "0.0%"}

# ─── ダッシュボード ───
sid = 0
requests.append(cell_fmt(sid, 0, 1, 0, 11, bg=DARK_BLUE, bold=True, fg=WHITE, font_size=14))
requests.append(cell_fmt(sid, 2, 3, 0, 11, bg=MID_BLUE, bold=True, fg=WHITE))
requests.append(cell_fmt(sid, 3, 3 + len(PJ_LIST) + 2, 4, 9, num_fmt=YEN_FMT))
requests.append(cell_fmt(sid, 3, 3 + len(PJ_LIST) + 2, 9, 10, num_fmt=PCT_FMT))
for i in range(len(PJ_LIST)):
    bg = LIGHT_GRAY if i % 2 == 0 else WHITE
    requests.append(cell_fmt(sid, 3+i, 4+i, 0, 11, bg=bg))
# 合計行
requests.append(cell_fmt(sid, 3+len(PJ_LIST)+1, 3+len(PJ_LIST)+2, 0, 11, bg=LIGHT_YELLOW, bold=True))
for c, w in [(0,90),(1,150),(2,230),(3,100),(4,120),(5,110),(6,110),(7,110),(8,110),(9,80),(10,100)]:
    requests.append(col_w(sid, c, w))
requests.append(freeze(sid, rows=3))

# ─── 個別PLシート ───
for pj in PJ_LIST:
    key = pj["key"]
    sid = pl_sheet_map[key][1]
    N = len(PL_SHEETS[key]["month_labels"])

    # タイトル行
    requests.append(cell_fmt(sid, R_TITLE, R_TITLE+1, 0, C_START+N+1, bg=DARK_BLUE, bold=True, fg=WHITE, font_size=12))
    # 基本情報ラベル
    requests.append(cell_fmt(sid, R_INFO, R_INFO+3, 0, 1, bold=True))
    # 基本情報の金額セル
    requests.append(cell_fmt(sid, R_INFO+1, R_INFO+2, 1, 2, num_fmt=YEN_FMT))
    requests.append(cell_fmt(sid, R_INFO+1, R_INFO+2, 4, 5, num_fmt=YEN_FMT))
    # サマリー
    requests.append(cell_fmt(sid, R_SUMMARY_H, R_SUMMARY_H+1, 0, 6, bg=LIGHT_GREEN, bold=True))
    requests.append(cell_fmt(sid, R_SUMMARY_V, R_SUMMARY_V+1, 1, 4, num_fmt=YEN_FMT))
    requests.append(cell_fmt(sid, R_SUMMARY_V, R_SUMMARY_V+1, 4, 5, num_fmt=PCT_FMT))
    requests.append(cell_fmt(sid, R_SUMMARY_V, R_SUMMARY_V+1, 5, 6, num_fmt=YEN_FMT))
    # 月別ヘッダー
    requests.append(cell_fmt(sid, R_MONTHLY_H, R_MONTHLY_H+1, 0, C_START+N+1, bg=LIGHT_BLUE, bold=True))
    # 月別の金額セル
    requests.append(cell_fmt(sid, R_SALES, R_EXTERNAL+1, 2, C_START+N+1, num_fmt=YEN_FMT))
    requests.append(cell_fmt(sid, R_GROSS, R_GROSS+1, 2, C_START+N+1, num_fmt=YEN_FMT))
    requests.append(cell_fmt(sid, R_GROSS_PCT, R_GROSS_PCT+1, 2, C_START+N+1, num_fmt=PCT_FMT))
    # 粗利行の背景
    requests.append(cell_fmt(sid, R_GROSS, R_GROSS+1, 0, C_START+N+1, bg=LIGHT_YELLOW, bold=True))
    # 工数ヘッダー
    requests.append(cell_fmt(sid, R_LABOR_H, R_LABOR_H+1, 0, C_START+N+2, bg=LIGHT_BLUE, bold=True))
    # 工数の配分率（%表示）— D列以降（C列は単価）
    requests.append(cell_fmt(sid, R_LABOR_START, R_LABOR_START+MAX_MEMBERS, L_START, L_START+N, num_fmt=PCT_FMT))
    # 工数合計金額（合計列）
    requests.append(cell_fmt(sid, R_LABOR_START, R_LABOR_SUB+1, L_START+N, L_START+N+1, num_fmt=YEN_FMT))
    # 工数小計
    requests.append(cell_fmt(sid, R_LABOR_SUB, R_LABOR_SUB+1, 0, L_START+N+1, bg=LIGHT_GRAY, bold=True, num_fmt=YEN_FMT))
    # 単価列
    requests.append(cell_fmt(sid, R_LABOR_START, R_LABOR_START+MAX_MEMBERS, 2, 3, num_fmt=YEN_FMT))
    # 外注ヘッダー
    requests.append(cell_fmt(sid, R_VENDOR_H, R_VENDOR_H+1, 0, C_START+N+1, bg=LIGHT_BLUE, bold=True))
    # 外注金額
    requests.append(cell_fmt(sid, R_VENDOR_START, R_VENDOR_SUB+1, 2, C_START+N+1, num_fmt=YEN_FMT))
    # 外注小計
    requests.append(cell_fmt(sid, R_VENDOR_SUB, R_VENDOR_SUB+1, 0, C_START+N+1, bg=LIGHT_GRAY, bold=True))
    # 列幅
    requests.append(col_w(sid, 0, 100))
    requests.append(col_w(sid, 1, 120))
    for c in range(2, C_START+N+1):
        requests.append(col_w(sid, c, 110))
    requests.append(freeze(sid, rows=R_MONTHLY_H+1, cols=2))

# ─── 工数管理 ───
sid = 50
requests.append(cell_fmt(sid, 0, 1, 0, 3, bg=DARK_BLUE, bold=True, fg=WHITE, font_size=14))
requests.append(cell_fmt(sid, 2, 3, 0, 3, bg=MID_BLUE, bold=True, fg=WHITE))
requests.append(cell_fmt(sid, 3, 200, 2, 3, num_fmt=YEN_FMT))
requests.append(col_w(sid, 0, 160))
requests.append(col_w(sid, 1, 140))
requests.append(col_w(sid, 2, 120))
requests.append(freeze(sid, rows=3))

# ─── 外注費管理 ───
sid = 51
requests.append(cell_fmt(sid, 0, 1, 0, 3, bg=DARK_BLUE, bold=True, fg=WHITE, font_size=14))
requests.append(cell_fmt(sid, 2, 3, 0, 3, bg=MID_BLUE, bold=True, fg=WHITE))
requests.append(cell_fmt(sid, 3, 200, 2, 3, num_fmt=YEN_FMT))
requests.append(col_w(sid, 0, 160))
requests.append(col_w(sid, 1, 140))
requests.append(col_w(sid, 2, 120))
requests.append(freeze(sid, rows=3))

# ─── PJ台帳 ───
sid = 1
requests.append(cell_fmt(sid, 0, 1, 0, 9, bg=DARK_BLUE, bold=True, fg=WHITE, font_size=14))
requests.append(cell_fmt(sid, 2, 3, 0, 9, bg=MID_BLUE, bold=True, fg=WHITE))
requests.append(cell_fmt(sid, 3, 30, 4, 5, num_fmt=YEN_FMT))
requests.append(cell_fmt(sid, 3, 30, 8, 9, num_fmt=YEN_FMT))
for c, w in [(0,90),(1,150),(2,230),(3,100),(4,120),(5,110),(6,110),(7,80),(8,110)]:
    requests.append(col_w(sid, c, w))
requests.append(freeze(sid, rows=3))

# 一括適用
sheets_svc.spreadsheets().batchUpdate(
    spreadsheetId=NEW_SSID, body={"requests": requests},
).execute()
print("書式設定完了")

# 共有
drive_svc.permissions().create(
    fileId=NEW_SSID, body={"role": "writer", "type": "anyone"},
).execute()
print("共有設定完了")

print(f"\n✅ 完成！")
print(f"https://docs.google.com/spreadsheets/d/{NEW_SSID}")
