import requests
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from datetime import datetime, timedelta
import jpholiday
from gspread_formatting import *
import calendar

# -------------------------
# 設定
# -------------------------
SERVICE_ACCOUNT_FILE = 'service_account.json'  # サービスアカウントjson
API_TOKEN = '41eL_54-bynysLzAsmad'
API_BASE = 'https://site1.nyutai.com/api/chief/v1'
休校日マスター_SHEET_ID = '1jjAZQDcy1JSdnYqQECNEcOXnr8-2-aKp-70mGNQdM0E'
休校日マスター_SHEET_NAME = '休校日マスター'  # シート1

TARGET_MONTH = '2025-07'  # 対象月
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# -------------------------
# Google認証
# -------------------------
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
gc = gspread.authorize(creds)

# -------------------------
# 休校日マスター読み込み
# -------------------------
休校日_ws = gc.open_by_key(休校日マスター_SHEET_ID).worksheet(休校日マスター_SHEET_NAME)
休校日_dict = {}
休校日_raw = 休校日_ws.get_all_values()
for row in 休校日_raw[1:]:
    if row and row[0]:
        休校日_dict[row[0]] = row[1] if len(row) > 1 else ""

# -------------------------
# 生徒一覧取得
# -------------------------
HEADERS = {"Api-Token": API_TOKEN}
students_resp = requests.get(f"{API_BASE}/students", headers=HEADERS)
students = students_resp.json()['data']

# -------------------------
# 日付リスト作成
# -------------------------
start_date = datetime.strptime(TARGET_MONTH + "-01", "%Y-%m-%d")
end_date = (start_date.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
days = [(start_date + timedelta(days=i)).strftime("%Y-%m-%d") for i in range((end_date - start_date).days + 1)]
day_colnames = [f"{int(day[-2:])}日" for day in days]

# -------------------------
# ファイル作成
# -------------------------
file_name = f"{TARGET_MONTH.replace('-', '年')}月出席簿"
sh = gc.create(file_name)

# -------------------------
# 色設定
# -------------------------
color_saturday = CellFormat(backgroundColor=Color(1, 0.95, 0.8))  # 薄オレンジ
color_sunday = CellFormat(backgroundColor=Color(1, 0.8, 0.9))     # 薄ピンク
color_holiday = CellFormat(backgroundColor=Color(0.9, 0.95, 1))   # 薄青
color_school_closed = CellFormat(backgroundColor=Color(0.85, 0.85, 0.85)) # グレー

# -------------------------
# 各生徒ごとにシート作成
# -------------------------
for stu in students:
    name = stu['name']
    # 入退室記録取得
    params = {
        'user_id': stu['id'],
        'date_from': days[0],
        'date_to': days[-1],
        'sort_key': 'entrance_time',
        'sort_dir': 'asc'
    }
    recs = requests.get(f"{API_BASE}/entrance_and_exits", headers=HEADERS, params=params).json()['data']
    # 日ごとに最大3回まで記録
    day_records = {day: [] for day in days}
    for rec in recs:
        d = rec['entrance_time'][:10]
        t_in = rec['entrance_time'][11:16] if rec['entrance_time'] else ''
        t_out = rec['exit_time'][11:16] if rec['exit_time'] else ''
        if len(day_records[d]) < 3:
            day_records[d].append(f"{t_in}-{t_out}" if t_in and t_out else "-")
    # カレンダー用テーブル作成
    data = []
    for i in range(3):  # 1日最大3回
        row = []
        for j, day in enumerate(days):
            dt = datetime.strptime(day, "%Y-%m-%d")
            # 休校日理由優先
            if day in 休校日_dict:
                cell_content = f"休校（{休校日_dict[day]}）" if 休校日_dict[day] else "休校"
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
            # 1回でも登校があれば1カウント
            shusseki.append(1 if any(x not in ("-", "", None) for x in day_records[day]) else 0)
    # 合計と備考
    合計 = sum(shusseki)
    data.append(shusseki)
    data.append([""] * len(days))  # 備考行（空欄）

    # 列名: 日付＋月合計＋備考
    columns = day_colnames + ["月合計", "備考"]
    # DataFrame化（行：1回目、2回目、3回目、出席数、備考）
    df = pd.DataFrame(data, index=["1回目", "2回目", "3回目", "出席数", "備考"], columns=day_colnames)
    df["月合計"] = ["", "", "", 合計, ""]
    df["備考"] = [""] * 4 + [""]

    # シート作成＆転記
    ws = sh.add_worksheet(title=name, rows=str(10), cols=str(len(columns)+1))
    # ヘッダー＋データ行リスト化
    update_values = [[""] + columns] + [[df.index[i]] + list(df.iloc[i]) for i in range(len(df))]
    ws.update(update_values)

    # 色分け
    for col, day in enumerate(days, start=2):  # B列から
        dt = datetime.strptime(day, "%Y-%m-%d")
        cell_range = gspread.utils.rowcol_to_a1(2, col) + ":" + gspread.utils.rowcol_to_a1(5, col)
        if day in 休校日_dict:
            format_cell_range(ws, cell_range, color_school_closed)
        elif jpholiday.is_holiday(dt):
            format_cell_range(ws, cell_range, color_holiday)
        elif dt.weekday() == 5:
            format_cell_range(ws, cell_range, color_saturday)
        elif dt.weekday() == 6:
            format_cell_range(ws, cell_range, color_sunday)

# デフォルトの「Sheet1」削除
try:
    sh.del_worksheet(sh.worksheet("Sheet1"))
except Exception:
    pass

print(f"全生徒分の出席簿作成が完了しました！ファイル名: {file_name}")
