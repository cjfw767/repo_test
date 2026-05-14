# 설정 : Library
import os, time, re, json
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from pytz import timezone

import gspread
from gspread_dataframe import set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials

# 설정 : 날짜
kst = timezone('Asia/Seoul')
today = datetime.now(kst)
base_date = today - timedelta(days=1)
bd_year, bd_month = base_date.year, base_date.month
pm_year, pm_month = (bd_year - 1, 12) if bd_month == 1 else (bd_year, bd_month - 1)

target_periods = [
    (pm_year, pm_month),
    (bd_year, bd_month),
]

# 설정 : 계정 정보
fs_json_path = os.getenv("FS_JSON_PATH", "fs_accounts.json")
gs_json_path = os.getenv("GS_JSON_PATH", "gs_accounts.json")
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]
creds = ServiceAccountCredentials.from_json_keyfile_name(gs_json_path, scope)
gc = gspread.authorize(creds)

spreadsheet_url = "https://docs.google.com/spreadsheets/d/1Rj05cusbuxkt1UTQXQgYIXVxncV-cT85zDIZrb7qYZw/edit"
spreadsheet = gc.open_by_url(spreadsheet_url)

# 식봄 판매자센터 접속
with open(fs_json_path, 'r') as f:
    pw_dict = json.load(f)

os.environ["FOODSPRING_PASS_CENTRAL"] = pw_dict.get("CENTRAL", "")
os.environ["FOODSPRING_PASS_WEST"] = pw_dict.get("WEST", "")
os.environ["FOODSPRING_PASS_EAST"] = pw_dict.get("EAST", "")

accounts = [
    ("Central", "platform@cj.net",   os.getenv("FOODSPRING_PASS_CENTRAL")),
    ("West",    "platformjs@cj.net", os.getenv("FOODSPRING_PASS_WEST")),
    ("East",    "platformys@cj.net", os.getenv("FOODSPRING_PASS_EAST")),
]

# 실적 정보 수집
def collect_order_stats(account_name, user_id, user_pw):
    print(f"수집 시작 : {account_name}")

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    driver = webdriver.Chrome(options=options)

    results, headers = [], []

    try:
        driver.get("https://seller.foodspring.co.kr/")
        time.sleep(3)
        driver.find_element(By.ID, "vendor_id").send_keys(user_id)
        driver.find_element(By.ID, "vendor_pass").send_keys(user_pw)
        driver.find_element(By.ID, "login-btn").click()
        time.sleep(5)

        for year, month in target_periods:
            url = f"https://seller.foodspring.co.kr/vendor/stats/stats_order?chk_search=Y&syear={year}&smonth={month}"
            driver.get(url)
            time.sleep(10)

            soup = BeautifulSoup(driver.page_source, 'html.parser')
            table = soup.find("table", id="stats_table")
            if not table:
                print(f"정보 없음 : {account_name} - {year}-{month}")
                continue

            if not headers:
                headers = [th.get_text(strip=True) for th in table.find("thead").find_all("th")]
                headers.append("Account")

            for row in table.find("tbody").find_all("tr"):
                cols = [td.get_text(strip=True) for td in row.find_all("td")]
                cleaned = []
                for col in cols:
                    val = re.sub(r"[건원,]", "", col)
                    try:
                        val = int(val)
                    except ValueError:
                        pass
                    cleaned.append(val)
                cleaned.append(account_name)
                results.append(cleaned)

    finally:
        driver.quit()
        print(f"수집 완료 : {account_name}")

    return results, headers

# 실적 정보 저장 (Google Spreadsheet)
all_data = []
header = []

for account_name, user_id, user_pw in accounts:
    data, current_header = collect_order_stats(account_name, user_id, user_pw)
    all_data.extend(data)
    if not header and current_header:
        header = current_header

df = pd.DataFrame(all_data, columns=header)
df["기간"] = pd.to_datetime(df["기간"], format="%Y.%m.%d", errors="coerce").dt.strftime("%Y-%m-%d")

for account_name in df["Account"].unique():
    df_account = df[df["Account"] == account_name].drop(columns="Account")

    try:
        try:
            spreadsheet.del_worksheet(spreadsheet.worksheet(account_name))
        except:
            pass

        worksheet = spreadsheet.add_worksheet(title=account_name, rows="1000", cols="20")
        set_with_dataframe(worksheet, df_account)
        print(f"저장 완료 : {account_name}")

    except Exception as e:
        print(f"저장 실패 : {account_name} / 오류 : {e}")

print("전체 작업 완료!")
