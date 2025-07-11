import requests
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, timedelta
import jpholiday
from gspread_formatting import *
import calendar
import re

# -----------------------------
# ★ ここをあなたの環境に合わせて設定してください ★
# -----------------------------
SERVICE_ACCOUNT_FILE = 'service_account.json'  # サービスアカウントjsonのファイル名
API_TOKEN = '41eL_54-bynysLzAsmad'
API_BASE = 'https://site1.nyutai.com/api/chief/v1'

# 休校日マスターのGoogleスプレッドシートURLからIDを抜き出してください
# 例：https://docs.google.com/spreadsheets/d/1jjAZQDcy1JSdnYqQECNEcOXnr8-2-aKp-70mGNQdM0E/edit#gid=0
休校日マスター_SHEET_ID = '1jjAZQDcy1JSdnYqQECNEcOXnr8-2-aKp-70mGNQdM0E'
休校日マスター_SHEET_NAME = '休校日マスター'  # シート名（通常 "休校日マスター"）

TARGET_MONTH = '2025-07'  # 生成したい年月（例：'2025-07'）

# -----------------------------
# Google認証
# -----------------------------
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
gc = gspread.authorize(creds)

# -----------------------------
# 休校日マスターを取得
# -----------------------------
休校日_ws = gc.open_by_key(休校日マスター_SHEET_ID).worksheet(休校日マスター_SHEET_NAME)
休校日_dict = {}
rows = 休校日_ws.get_all_values()
for row in rows[1:]:  # 1行目はヘッダー
    if row and row[0]:
        if re.match(r'^\d{4}-\d{2}-\d{2}$', row[0]):
            休校日_dict[row[0]] = row[1] if len(row) > 1 else ""

# -----------------------------
# 生徒一覧をAPIで取得
# -----------------------------
HEADERS = {"Api-Token": API_TOKEN}
students_resp = requests.get(f"{API_BASE}/students", headers=HEADERS)
students = students_resp.json()['data']

# -----------------------------
# 月の日付リスト作成
# -----------------------------
start_date = datetime.strptime(TARGET_MONTH + "-01", "%Y-%m-%d")
end_date = (start_date.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
days = [(start_date + timedelta(days=i)).strftime("%Y-%m-%d") for i in range((end_date - start_date).days + 1)]
day_colnames = [f"{int(day[-2:])}日" for day in days]

# -----------------------------
# 出席簿ファイル作成
# -----------------------------
file_name = f"{TARGET_MONTH.replace('-', '年')}月出席簿"
sh = gc.create(file_name)

# -----------------------------
# 色設定
# -----------------------------
color_saturday = CellFormat(backgroundColor=Color(1, 0.95, 0.8))  # 薄オレンジ
color_sunday = CellFormat(backgroundColor=Color(1, 0.8, 0.9))     # 薄ピンク
color_holiday = CellFormat(backgroundColor=Color(0.9, 0.95, 1))   # 薄青
color_school_closed = CellFormat(backgroundColor=Color(0.85, 0.85, 0.85)) # グレー

# -----------------------------
# 生徒ごとにシートを作成＆書き込み
# -----------------------------
for stu in students:
    name = stu['name']
    # 入退室記録をAPIから取得
    params = {
        'user_id': stu['id'],
        'date_from': days[0],
        'date_to': days[-1],
        'sort_key': 'entrance_time',
        'sort_dir': 'asc'
    }
    recs = requests.get(f"{API_BASE}/entrance_and_exits", headers=HEADERS, params=params).json()['data']
    # 日ごと最大3回分
    day_records = {day: [] for day in days}
    for rec in recs:
        d = rec['entrance_time'][:10]
        t_in = rec['entrance_time'][11:16] if rec['entrance_time'] else ''
        t_out = rec['exit_time'][11:16] if rec['exit_time'] else ''
        if len(day_records[d]) < 3:
            day_records[d].append(f"{t_in}-{t_out}" if t_in and t_out else "-")
    # テーブル作成
    data = []
    for i in range(3):
        row = []
        for j, day in enumerate(days):
            dt = datetime.strptime(day, "%Y-%m-%d")
            if day in 休校日_dict:
                reason = 休校日_dict[day]
                cell_content = f"休校（{reason}）" if reason else "休校"
            elif jpholiday.is_holiday_name(dt):
                cell_content = f"祝日（{jpholiday.is_holiday_name(dt)}）"
            elif dt.weekday() == 5:
                cell_content = "土曜"
            elif dt.weekday() == 6:
                cell_content = "日曜"
            else:
                cell_content = day_records[day][i] if len(day_records[day]) > i else "-"
            row.append(cell_content)
        data.append(row)
    # 出席数行
    shusseki = []
    for j, day in enumerate(days):
        dt = datetime.strptime(day, "%Y-%m-%d")
        if day in 休校日_dict or jpholiday.is_holiday(dt) or dt.weekday() >= 5:
            shusseki.append(0)
        else:
            shusseki.append(1 if any(x not in ("-", "", None) for x in day_records[day]) else 0)
    合計 = sum(shusseki)
    data.append(shusseki)
    data.append([""] * len(days))  # 備考

    # 列名
    columns = day_colnames + ["月合計", "備考"]
    df = pd.DataFrame(data, index=["1回目", "2回目", "3回目", "出席数", "備考"], columns=day_colnames)
    df["月合計"] = ["", "", "", 合計, ""]
    df["備考"] = [""] * 4 + [""]

    # シート作成
    ws = sh.add_worksheet(title=name, rows=str(10), cols=str(len(columns)+1))
    update_values = [[""] + columns] + [[df.index[i]] + list(df.iloc[i]) for i in range(len(df))]
    ws.update(update_values)

    # -----------------------------
    # 一括で色塗りするロジック（APIリクエスト大幅削減！）
    # -----------------------------
    color_cols = defaultdict(list)
    for col, day in enumerate(days, start=2):  # B列から
        dt = datetime.strptime(day, "%Y-%m-%d")
        if day in 休校日_dict:
            color_cols["school_closed"].append(col)
        elif jpholiday.is_holiday(dt):
            color_cols["holiday"].append(col)
        elif dt.weekday() == 5:
            color_cols["saturday"].append(col)
        elif dt.weekday() == 6:
            color_cols["sunday"].append(col)
    # 範囲を連続区間ごとにまとめて色塗り
    def group_ranges(cols):
        if not cols:
            return []
        cols = sorted(cols)
        groups = [[cols[0]]]
        for c in cols[1:]:
            if c == groups[-1][-1] + 1:
                groups[-1].append(c)
            else:
                groups.append([c])
        return groups

    color_map = {
        "school_closed": color_school_closed,
        "holiday": color_holiday,
        "saturday": color_saturday,
        "sunday": color_sunday
    }

    for key in color_cols:
        for group in group_ranges(color_cols[key]):
            start_col = group[0]
            end_col = group[-1]
            rng = f"{gspread.utils.rowcol_to_a1(2, start_col)}:{gspread.utils.rowcol_to_a1(5, end_col)}"
            format_cell_range(ws, rng, color_map[key])

# デフォルトのSheet1削除
try:
    sh.del_worksheet(sh.worksheet("Sheet1"))
except Exception:
    pass

print(f"全生徒分の出席簿作成が完了しました！ファイル名: {file_name}")
