import streamlit as st
import pandas as pd
import cloudscraper
from bs4 import BeautifulSoup
import re
import os
import glob
from datetime import datetime, timedelta

# ==========================================
# 網頁基本設定 (必須放在最前面)
# ==========================================
st.set_page_config(page_title="總經日曆儀表板", layout="wide", page_icon="📊")

# ==========================================
# 模組 1：爬蟲核心功能
# ==========================================
def fetch_and_save_data(start_date, end_date):
    scraper = cloudscraper.create_scraper(browser={
        'browser': 'chrome', 'platform': 'windows', 'desktop': True
    })
    api_url = "https://hk.investing.com/economic-calendar/Service/getCalendarFilteredData"
    
    headers = {
        'X-Requested-With': 'XMLHttpRequest',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Referer': 'https://hk.investing.com/economic-calendar/',
        'Origin': 'https://hk.investing.com'
    }
    
    # 5: 美國, 35: 日本, 72: 歐元區
    payload = {
        'country[]': ['5', '35', '72'], 
        'dateFrom': start_date, 'dateTo': end_date,
        'timeZone': '8', 'timeFilter': 'timeRemain', 'currentTab': 'custom', 'limit_from': '0'
    }
    
    response = scraper.post(api_url, headers=headers, data=payload)
    if response.status_code != 200:
        return None, f"請求失敗，狀態碼: {response.status_code}"
        
    try:
        html_data = response.json().get('data', '')
        soup = BeautifulSoup(html_data, 'html.parser')
        events = []
        current_date_str = ""
        
        for row in soup.find_all('tr'):
            # 處理日期
            date_td = row.find('td', class_='theDay')
            if date_td:
                raw_date = date_td.text.strip()
                nums = re.findall(r'\d+', raw_date)
                if len(nums) >= 3:
                    if len(nums[0]) == 4:
                        current_date_str = f"{nums[0]}-{int(nums[1]):02d}-{int(nums[2]):02d}"
                    elif len(nums[2]) == 4:
                        current_date_str = f"{nums[2]}-{int(nums[0]):02d}-{int(nums[1]):02d}"
                    else:
                        current_date_str = raw_date
                else:
                    current_date_str = raw_date
                continue
            
            # 處理事件
            if 'js-event-item' in row.get('class', []):
                event_td = row.find('td', class_='event')
                event_name = event_td.text.strip() if event_td else ''
                if not event_name: continue
                
                flag_span = row.find('span', class_='ceFlags')
                country = flag_span.get('title') if flag_span else '未知'
                
                actual = row.find('td', class_='act').text.strip() if row.find('td', class_='act') else ''
                forecast = row.find('td', class_='fore').text.strip() if row.find('td', class_='fore') else ''
                previous = row.find('td', class_='prev').text.strip() if row.find('td', class_='prev') else ''
                
                events.append({
                    '日期': current_date_str, '國家': country, '事件': event_name, 
                    '今值': actual, '預測': forecast, '前值': previous
                })
        
        df = pd.DataFrame(events)
        
        # 存檔為 Excel
        filename = f'精簡版_經濟日曆_{start_date}_至_{end_date}.xlsx'
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            for c in df['國家'].unique():
                sheet_name = str(c)[:31] if str(c) else '其他'
                df[df['國家']==c].to_excel(writer, sheet_name=sheet_name, index=False)
        return filename, None
        
    except Exception as e:
        return None, f"解析過程發生錯誤: {str(e)}"

# ==========================================
# 模組 2：資料讀取與快取機制
# ==========================================
@st.cache_data
def load_latest_data():
    """自動尋找桌面最新抓取的精簡版 Excel 並讀取"""
    files = glob.glob('精簡版_經濟日曆*.xlsx')
    if not files:
        return pd.DataFrame(), None
        
    # 找到修改時間最新的檔案
    latest_file = max(files, key=os.path.getmtime)
    
    # 將所有 Sheet 合併成一個 DataFrame 以供儀表板篩選
    xls = pd.read_excel(latest_file, sheet_name=None)
    df_list = [df_sheet for sheet_name, df_sheet in xls.items()]
    df = pd.concat(df_list, ignore_index=True)
    return df, latest_file

# ==========================================
# 模組 3：互動式儀表板 (前端 UI)
# ==========================================
st.title("📊 總體經濟數據儀表板")
st.markdown("一站式獲取並分析美國、日本、歐元區的重要總經數據。")

# --- 側邊欄：抓取控制區 ---
st.sidebar.header("🔄 1. 獲取最新數據")
today = datetime.now()
default_start = today - timedelta(days=today.weekday())
default_end = today + timedelta(days=11 - today.weekday())

scrape_start = st.sidebar.date_input("選擇開始日期", default_start)
scrape_end = st.sidebar.date_input("選擇結束日期", default_end)

if st.sidebar.button("🚀 立即啟動爬蟲", use_container_width=True):
    with st.spinner("正在向 Investing.com 抓取數據中..."):
        fname, err = fetch_and_save_data(scrape_start.strftime('%Y-%m-%d'), scrape_end.strftime('%Y-%m-%d'))
        if fname:
            st.sidebar.success("✅ 數據更新完成！")
            st.cache_data.clear() # 🌟 關鍵優化：清除舊緩存，強迫讀取新檔案
            st.rerun() # 重新整理網頁
        else:
            st.sidebar.error(f"❌ 抓取失敗：{err}")

st.sidebar.divider()

# --- 載入資料 ---
df, current_file = load_latest_data()

# --- 側邊欄：篩選控制區 ---
if not df.empty:
    st.sidebar.header("🔍 2. 篩選與下載")
    st.sidebar.info(f"📂 目前讀取：\n**{current_file}**")
    
    # 實體下載按鈕
    with open(current_file, "rb") as file:
        st.sidebar.download_button(
            label="📥 下載此 Excel 報表",
            data=file,
            file_name=current_file,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    
    # 國家篩選
    countries = df['國家'].dropna().unique().tolist()
    selected_countries = st.sidebar.multiselect("選擇國家 / 地區", options=countries, default=countries)
    
    # 日期篩選
    df['暫存日期格式'] = pd.to_datetime(df['日期'], errors='coerce')
    valid_dates = df['暫存日期格式'].dropna()
    if not valid_dates.empty:
        min_date, max_date = valid_dates.min().date(), valid_dates.max().date()
        date_range = st.sidebar.date_input("檢視日期範圍", [min_date, max_date])
    else:
        date_range = []

    # 根據篩選器過濾資料
    filtered_df = df[df['國家'].isin(selected_countries)]
    if len(date_range) == 2:
        start_date, end_date = date_range
        filtered_df = filtered_df[(filtered_df['暫存日期格式'].dt.date >= start_date) & 
                                  (filtered_df['暫存日期格式'].dt.date <= end_date)]

    # 移除暫存運算欄位
    filtered_df = filtered_df.drop(columns=['暫存日期格式'], errors='ignore')

    # 主畫面顯示
    st.subheader(f"共找到 {len(filtered_df)} 筆經濟數據")
    st.dataframe(filtered_df, use_container_width=True, hide_index=True)

else:
    # 找不到資料時的友善提示
    st.info("👋 歡迎！目前系統找不到已下載的經濟數據。")
    st.write("👉 請使用左側面板設定日期，並點擊 **「🚀 立即啟動爬蟲」** 來抓取第一份數據。")