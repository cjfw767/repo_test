# 설정 : Library
import pandas as pd
import datetime
import gspread
from google.oauth2.service_account import Credentials
import requests
import time
import os
from pytz import timezone

# 설정 : 계정 정보
GS_JSON_PATH = os.getenv("GS_JSON_PATH", "gs_accounts.json")
scope = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
creds = Credentials.from_service_account_file(GS_JSON_PATH, scopes=scope)
gc = gspread.authorize(creds)

spreadsheet = gc.open_by_url("https://docs.google.com/spreadsheets/d/12jkPCRuMfDu26pDy6mg3_gMUsEl2N60yGkBscJsIchM/edit?usp=sharing")
sheet_names = ["Central", "West", "East"]
region_map = {"Central": "중부", "East": "동부", "West": "서부"}

# 설정 : 날짜
kst = timezone('Asia/Seoul')
today = datetime.datetime.now(kst)
yesterday = today - datetime.timedelta(days=1)
yesterday_str = yesterday.strftime('%Y-%m-%d')
last_week = yesterday - datetime.timedelta(days=7)
last_week_str = last_week.strftime('%Y-%m-%d')

month_start = yesterday.replace(day=1)
prev_month_end = month_start - datetime.timedelta(days=1)
prev_month_start = prev_month_end.replace(day=1)

date_range = pd.date_range(start=prev_month_start, end=yesterday)

summary_list = []
summary_list_prev_month = []
summary_list_this_month = []

yesterday_data = {}
last_week_data = {}

# 일별 자료 산정
for single_date in date_range:
    date_str = single_date.strftime('%Y-%m-%d')
    total_gmv = 0
    total_coupon = 0
    region_gmv = {}
    region_coupon = {}

    for sheet in sheet_names:
        worksheet = spreadsheet.worksheet(sheet)
        data = worksheet.get_all_records()
        time.sleep(5)

        df = pd.DataFrame(data)
        if '기간' not in df.columns:
            continue
        df['기간'] = pd.to_datetime(df['기간'], errors='coerce')
        day_df = df[df['기간'].dt.date == single_date.date()]
        if day_df.empty:
            continue

        region = region_map[sheet]
        gmv = day_df['주문금액'].sum()
        coupon = day_df['판매자쿠폰 사용금액'].sum()
        total_gmv += gmv
        total_coupon += coupon
        region_gmv[region] = gmv
        region_coupon[region] = coupon

    if total_gmv == 0 and total_coupon == 0:
        continue

    summary_data = {
        "날짜": date_str,
        "GMV_합계": total_gmv,
        "GMV_중부": region_gmv.get("중부", 0),
        "GMV_동부": region_gmv.get("동부", 0),
        "GMV_서부": region_gmv.get("서부", 0),
        "쿠폰_합계": total_coupon,
        "쿠폰_중부": region_coupon.get("중부", 0),
        "쿠폰_동부": region_coupon.get("동부", 0),
        "쿠폰_서부": region_coupon.get("서부", 0),
        "비용율_합계": round(total_coupon / total_gmv * 100, 2) / 100 if total_gmv else 0,
        "비용율_중부": round(region_coupon.get("중부", 0) / region_gmv.get("중부", 1) * 100, 2) / 100 if region_gmv.get("중부", 0) else 0,
        "비용율_동부": round(region_coupon.get("동부", 0) / region_gmv.get("동부", 1) * 100, 2) / 100 if region_gmv.get("동부", 0) else 0,
        "비용율_서부": round(region_coupon.get("서부", 0) / region_gmv.get("서부", 1) * 100, 2) / 100 if region_gmv.get("서부", 0) else 0
    }

    summary_list.append(summary_data)

    if single_date < month_start:
        summary_list_prev_month.append(summary_data)
    else:
        summary_list_this_month.append(summary_data)

    if date_str == yesterday_str:
        yesterday_data = summary_data
    elif date_str == last_week_str:
        last_week_data = summary_data

# 합계 자료 산정
def make_summary_row(label, data_list):
    gmv_all = sum(item['GMV_합계'] for item in data_list)
    coupon_all = sum(item['쿠폰_합계'] for item in data_list)
    region_gmv = {k: sum(item[f'GMV_{k}'] for item in data_list) for k in ['중부', '동부', '서부']}
    region_coupon = {k: sum(item[f'쿠폰_{k}'] for item in data_list) for k in ['중부', '동부', '서부']}
    return {
        "날짜": label,
        "GMV_합계": gmv_all,
        "GMV_중부": region_gmv["중부"],
        "GMV_동부": region_gmv["동부"],
        "GMV_서부": region_gmv["서부"],
        "쿠폰_합계": coupon_all,
        "쿠폰_중부": region_coupon["중부"],
        "쿠폰_동부": region_coupon["동부"],
        "쿠폰_서부": region_coupon["서부"],
        "비용율_합계": round(coupon_all / gmv_all * 100, 2) / 100 if gmv_all else 0,
        "비용율_중부": round(region_coupon["중부"] / region_gmv["중부"] * 100, 2) / 100 if region_gmv["중부"] else 0,
        "비용율_동부": round(region_coupon["동부"] / region_gmv["동부"] * 100, 2) / 100 if region_gmv["동부"] else 0,
        "비용율_서부": round(region_coupon["서부"] / region_gmv["서부"] * 100, 2) / 100 if region_gmv["서부"] else 0
    }

summary_list.append(make_summary_row("전월합계", summary_list_prev_month))
summary_list.append(make_summary_row("당월합계", summary_list_this_month))

summary_df = pd.DataFrame(summary_list)

# 저장 : Google Spreadsheet
result_spreadsheet = gc.open_by_url("https://docs.google.com/spreadsheets/d/12jkPCRuMfDu26pDy6mg3_gMUsEl2N60yGkBscJsIchM/edit?usp=sharing")
try:
    ws = result_spreadsheet.worksheet("일별실적")
    ws.batch_clear(['A2:Z1000'])
except gspread.exceptions.WorksheetNotFound:
    ws = result_spreadsheet.add_worksheet(title="일별실적", rows=str(len(summary_df)+1), cols=str(len(summary_df.columns)))

ws.update([summary_df.columns.tolist()] + summary_df.values.tolist())

# 메시지 발송 (팀즈 워크플로)
def format_money(val):
    return f"{int(val):,}원"

msg = f"""<strong>식봄 CJFW GMV</strong><br><br>
<b>[{yesterday_str}]</b><br>
 - GMV : {format_money(yesterday_data['GMV_합계'])}  ( 중부 : {format_money(yesterday_data['GMV_중부'])} | 동부 : {format_money(yesterday_data['GMV_동부'])} | 서부 : {format_money(yesterday_data['GMV_서부'])} )<br>
&nbsp;&nbsp;&nbsp;&nbsp;전주 : {format_money(last_week_data.get('GMV_합계', 0))}  ( 중부 : {format_money(last_week_data.get('GMV_중부', 0))} | 동부 : {format_money(last_week_data.get('GMV_동부', 0))} | 서부 : {format_money(last_week_data.get('GMV_서부', 0))} )<br>
 - 쿠폰비용율 : {yesterday_data['비용율_합계']:.2%}  ( 중부 {yesterday_data['비용율_중부']:.2%} | 동부 {yesterday_data['비용율_동부']:.2%} | 서부 {yesterday_data['비용율_서부']:.2%} )<br><br>

<b>[{month_start.month}월 누적]</b><br>
 - GMV : {format_money(make_summary_row('당월합계', summary_list_this_month)['GMV_합계'])}  ( 중부 : {format_money(make_summary_row('당월합계', summary_list_this_month)['GMV_중부'])} | 동부 : {format_money(make_summary_row('당월합계', summary_list_this_month)['GMV_동부'])} | 서부 : {format_money(make_summary_row('당월합계', summary_list_this_month)['GMV_서부'])} )<br>
 - 쿠폰비용율 : {make_summary_row('당월합계', summary_list_this_month)['비용율_합계']:.2%}  ( 중부 {make_summary_row('당월합계', summary_list_this_month)['비용율_중부']:.2%} | 동부 {make_summary_row('당월합계', summary_list_this_month)['비용율_동부']:.2%} | 서부 {make_summary_row('당월합계', summary_list_this_month)['비용율_서부']:.2%} )<br><br>

<a href=\"https://docs.google.com/spreadsheets/d/12jkPCRuMfDu26pDy6mg3_gMUsEl2N60yGkBscJsIchM\" target=\"_blank\">[일자별 실적]</a>
"""

teams_webhook_url = os.getenv(
    "TEAMS_WEBHOOK_URL",
    "https://defaultee6af5c5684f45399eb664793af080.27.environment.api.powerplatform.com:443/powerautomate/automations/direct/workflows/64650c7863ed42d2b202522ebb109d8a/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=4DTnGL8xef1wQti0Jr96ZzJiLC_j3XmGhLcyKP5c2ew",
)

try:
    response = requests.post(teams_webhook_url, json={"message": msg})
    if response.status_code in (200, 202):
        print("Teams 워크플로 메시지 전송 완료!")
    else:
        print(f"Teams 워크플로 호출 실패: {response.status_code} - {response.text}")
except Exception as e:
    print(f"Teams 워크플로 메시지 전송 중 오류 발생: {e}")

print("전체 작업 완료!")
