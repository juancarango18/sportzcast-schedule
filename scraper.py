import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta
import re
import time
import os
import json

from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# Setup Chrome options for the Cloud
chrome_options = Options()
chrome_options.add_argument("--headless") # Runs Chrome in invisible mode
chrome_options.add_argument("--no-sandbox") # Bypasses OS security model
chrome_options.add_argument("--disable-dev-shm-usage") # Overcomes limited resource problems

# Initialize the driver
driver = webdriver.Chrome(options=chrome_options)

# ==========================================
# 1. DYNAMIC CONFIGURATION
# ==========================================

TARGET_YEAR = 2026
if os.path.exists("ui_inputs.json"):
    with open("ui_inputs.json", "r") as f:
        ui_data = json.load(f)
        TARGET_YEAR = ui_data.get("TARGET_YEAR", 2026)

WNBA_URL = f"https://www.wnba.com/schedule?season={TARGET_YEAR}&month=all"
CFL_URL = f"https://www.cfl.ca/schedule/{TARGET_YEAR}/"

HEADERS = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'
}

scraped_games = []

def convert_to_military_time(time_str):
    try:
        clean_time = time_str.strip().replace('.', '').upper()
        start_dt = datetime.strptime(clean_time, "%I:%M %p")
        end_dt = start_dt + timedelta(hours=3)
        return start_dt.strftime("%H:%M"), end_dt.strftime("%H:%M")
    except Exception:
        return None, None

def get_selenium_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new") 
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument(f"user-agent={HEADERS['User-Agent']}")
    chrome_options.add_argument("--log-level=3") 
    return webdriver.Chrome(options=chrome_options)

def scrape_wnba():
    print(f"🏀 Loading WNBA {TARGET_YEAR} schedule...")
    try:
        driver = get_selenium_driver()
        driver.get(WNBA_URL)
        time.sleep(5) 
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        driver.quit() 
        
        time_tags = soup.find_all('time', datetime=True)
        games_found = 0
        for time_tag in time_tags:
            try:
                dt_str = time_tag['datetime']
                utc_time = datetime.strptime(dt_str, "%Y-%m-%dT%H:%M:%SZ")
                local_time = utc_time - timedelta(hours=5) 
                
                date_text = local_time.strftime("%Y-%m-%d")
                start_time = local_time.strftime("%H:%M")
                
                card = time_tag.parent
                teams = []
                while card:
                    teams = card.find_all('p', class_=re.compile(r'_TeamName__name'))
                    if len(teams) >= 2: break 
                    card = card.parent
                
                if len(teams) >= 2:
                    scraped_games.append({
                        "Date": date_text, "Sport": "WNBA",
                        "Matchup": f"{teams[0].text.strip()} vs. {teams[1].text.strip()}",
                        "Coverage_Start": start_time, "Coverage_End": ""
                    })
                    games_found += 1
            except Exception:
                continue
        print(f"  ✅ Extracted {games_found} WNBA games!")
    except Exception as e:
        print(f"  ❌ Error on WNBA: {e}")

def scrape_cfl():
    print(f"🏈 Loading CFL {TARGET_YEAR} schedule...")
    try:
        driver = get_selenium_driver()
        driver.get(CFL_URL)
        time.sleep(5) 
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        driver.quit()
        
        date_time_divs = soup.find_all('div', class_='date-time')
        games_found = 0
        for dt_div in date_time_divs:
            try:
                date_span = dt_div.find('span', class_='date')
                time_span = dt_div.find('span', class_='time')
                if not date_span or not time_span: continue

                raw_date = date_span.text.strip()
                clean_date_str = f"{raw_date} {TARGET_YEAR}"
                date_obj = datetime.strptime(clean_date_str, "%a %b %d %Y")
                date_text = date_obj.strftime("%Y-%m-%d")

                raw_time = time_span.text.strip() 
                clean_time = raw_time.split('-')[0].split('+')[0].strip()
                start_time, end_time = convert_to_military_time(clean_time)
                if not start_time: continue

                matchup_div = dt_div.parent.find('div', class_='matchup')
                if matchup_div:
                    visitor_span = matchup_div.find('span', class_='visitor').find('span', class_='text')
                    host_span = matchup_div.find('span', class_='host').find('span', class_='text')
                    away_team = visitor_span.text.strip() if visitor_span else "Away"
                    home_team = host_span.text.strip() if host_span else "Home"

                    scraped_games.append({
                        "Date": date_text, "Sport": "CFL",
                        "Matchup": f"{away_team} vs. {home_team}",
                        "Coverage_Start": start_time, "Coverage_End": ""
                    })
                    games_found += 1
            except Exception:
                continue
        print(f"  ✅ Extracted {games_found} CFL games!")
    except Exception as e:
        print(f"  ❌ Error on CFL: {e}")

print("🚀 Starting Web Scraper...")
scrape_wnba()
scrape_cfl()

if scraped_games:
    df = pd.DataFrame(scraped_games)
    df = df.sort_values(by=['Date', 'Coverage_Start'])
    df.to_csv("games_schedule.csv", index=False)
    print("✅ SUCCESS! Live data saved.")
