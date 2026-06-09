import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta
import re
import time
import os
import json
import sys
import shutil
import traceback

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

def log(message):
    print(message, flush=True)

def log_environment():
    log("=== Scraper environment ===")
    log(f"  Python: {sys.executable}")
    log(f"  CWD: {os.getcwd()}")
    for var in ("CHROME_BIN", "CHROMEDRIVER_PATH", "PATH"):
        value = os.environ.get(var, "(not set)")
        if var == "PATH" and value != "(not set)":
            value = value[:200] + ("..." if len(value) > 200 else "")
        log(f"  {var}={value}")
    for label, path in [
        ("CHROME_BIN", os.environ.get("CHROME_BIN", "")),
        ("CHROMEDRIVER_PATH", os.environ.get("CHROMEDRIVER_PATH", "")),
        ("/usr/bin/chromium", "/usr/bin/chromium"),
        ("/usr/bin/chromium-browser", "/usr/bin/chromium-browser"),
        ("/usr/bin/chromedriver", "/usr/bin/chromedriver"),
    ]:
        if path:
            log(f"  {label} exists: {os.path.isfile(path)} ({path})")
    log(f"  shutil.which('chromium'): {shutil.which('chromium')}")
    log(f"  shutil.which('chromedriver'): {shutil.which('chromedriver')}")

# ==========================================
# 1. DYNAMIC CONFIGURATION & MEMORY BANKS
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

# The Venue Memory Banks!
WNBA_VENUES = {
    "Atlanta Dream": "Gateway Center Arena",
    "Chicago Sky": "Wintrust Arena",
    "Connecticut Sun": "Mohegan Sun Arena",
    "Dallas Wings": "College Park Center",
    "Indiana Fever": "Gainbridge Fieldhouse",
    "Las Vegas Aces": "Michelob ULTRA Arena",
    "Los Angeles Sparks": "Crypto.com Arena",
    "Minnesota Lynx": "Target Center",
    "New York Liberty": "Barclays Center",
    "Phoenix Mercury": "Mortgage Matchup Center",
    "Seattle Storm": "Climate Pledge Arena",
    "Washington Mystics": "CareFirst Arena",
    "Golden State Valkyries": "Chase Center",
    "Portland Fire": "Moda Center",
    "Toronto Tempo": "Coca-Cola Coliseum"
}

CFL_VENUES = {
    # Full Names
    "BC Lions": "BC Place",
    "Calgary Stampeders": "McMahon Stadium",
    "Edmonton Elks": "Commonwealth Stadium",
    "Saskatchewan Roughriders": "Mosaic Stadium",
    "Winnipeg Blue Bombers": "Princess Auto Stadium",
    "Hamilton Tiger-Cats": "Tim Hortons Field",
    "Toronto Argonauts": "BMO Field",
    "Ottawa Redblacks": "TD Place Stadium",
    "Montreal Alouettes": "Percival Molson Memorial Stadium",
    
    # Initials / Abbreviations
    "BC": "BC Place",
    "CGY": "McMahon Stadium",
    "EDM": "Commonwealth Stadium",
    "SSK": "Mosaic Stadium",
    "WPG": "Princess Auto Stadium",
    "HAM": "Tim Hortons Field",
    "TOR": "BMO Field",
    "OTT": "TD Place Stadium",
    "MTL": "Percival Molson Memorial Stadium"
}

scraped_games = []

def convert_to_military_time(time_str):
    try:
        # This regex looks specifically for the HH:MM AM/PM pattern and ignores " ET" or "EDT"
        match = re.search(r'(\d{1,2}:\d{2}\s*[aApP][mM])', time_str.replace('.', ''))
        if match:
            clean_time = match.group(1).upper()
            start_dt = datetime.strptime(clean_time, "%I:%M %p")
            end_dt = start_dt + timedelta(hours=3)
            return start_dt.strftime("%H:%M"), end_dt.strftime("%H:%M")
        return None, None
    except Exception:
        return None, None

def get_selenium_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument(f"user-agent={HEADERS['User-Agent']}")
    chrome_options.add_argument("--log-level=3")

    chrome_bin = os.environ.get("CHROME_BIN") or shutil.which("chromium") or "/usr/bin/chromium"
    if chrome_bin and os.path.isfile(chrome_bin):
        chrome_options.binary_location = chrome_bin
        log(f"  Using Chrome binary: {chrome_bin}")
    else:
        log(f"  WARNING: Chrome binary not found (tried CHROME_BIN, PATH, /usr/bin/chromium)")

    driver_path = os.environ.get("CHROMEDRIVER_PATH") or shutil.which("chromedriver") or "/usr/bin/chromedriver"
    if driver_path and os.path.isfile(driver_path):
        log(f"  Using ChromeDriver: {driver_path}")
        service = Service(executable_path=driver_path)
    else:
        log(f"  WARNING: ChromeDriver not found — letting Selenium auto-detect")
        service = Service()

    try:
        driver = webdriver.Chrome(service=service, options=chrome_options)
    except Exception:
        log("  FAILED to start Chrome/WebDriver:")
        log(traceback.format_exc())
        raise

    # Force the browser into Colombia time (UTC-5) for CFL schedule pages.
    try:
        driver.execute_cdp_cmd('Emulation.setTimezoneOverride', {
            'timezoneId': 'America/Bogota'
        })
    except Exception as e:
        log(f"  WARNING: could not set timezone override: {e}")

    return driver

def scrape_wnba():
    log(f"🏀 Loading WNBA {TARGET_YEAR} schedule from {WNBA_URL} ...")
    driver = None
    try:
        driver = get_selenium_driver()
        driver.get(WNBA_URL)
        time.sleep(5)
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        time_tags = soup.find_all('time', datetime=True)
        log(f"  Found {len(time_tags)} <time> elements on page")
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
                    if len(teams) >= 2:
                        break
                    card = card.parent

                if len(teams) >= 2:
                    away_team = teams[0].text.strip()
                    home_team = teams[1].text.strip()

                    venue = WNBA_VENUES.get(home_team, "TBD")

                    scraped_games.append({
                        "Date": date_text, "Sport": "WNBA",
                        "Matchup": f"{home_team} v {away_team}",
                        "Coverage_Start": start_time, "Coverage_End": "",
                        "Venue": venue
                    })
                    games_found += 1
            except Exception:
                continue
        log(f"  ✅ Extracted {games_found} WNBA games")
    except Exception as e:
        log(f"  ❌ WNBA scrape failed: {e}")
        log(traceback.format_exc())
        raise
    finally:
        if driver:
            driver.quit()

def scrape_cfl():
    log(f"🏈 Loading CFL {TARGET_YEAR} schedule from {CFL_URL} ...")
    driver = None
    try:
        driver = get_selenium_driver()
        driver.get(CFL_URL)
        time.sleep(5)
        soup = BeautifulSoup(driver.page_source, 'html.parser')

        date_time_divs = soup.find_all('div', class_='date-time')
        log(f"  Found {len(date_time_divs)} date-time blocks on page")
        games_found = 0
        for dt_div in date_time_divs:
            try:
                date_span = dt_div.find('span', class_='date')
                time_span = dt_div.find('span', class_='time')
                if not date_span or not time_span:
                    continue

                raw_date = date_span.text.strip()
                clean_date_str = f"{raw_date} {TARGET_YEAR}"
                date_obj = datetime.strptime(clean_date_str, "%a %b %d %Y")
                date_text = date_obj.strftime("%Y-%m-%d")

                raw_time = time_span.text.strip()
                clean_time = raw_time.split('-')[0].split('+')[0].strip()
                start_time, end_time = convert_to_military_time(clean_time)
                if not start_time:
                    continue

                matchup_div = dt_div.parent.find('div', class_='matchup')
                if matchup_div:
                    visitor_span = matchup_div.find('span', class_='visitor').find('span', class_='text')
                    host_span = matchup_div.find('span', class_='host').find('span', class_='text')

                    home_team = visitor_span.text.strip() if visitor_span else "Home"
                    away_team = host_span.text.strip() if host_span else "Away"

                    venue = CFL_VENUES.get(home_team, "TBD")

                    scraped_games.append({
                        "Date": date_text, "Sport": "CFL",
                        "Matchup": f"{away_team} v {home_team}",
                        "Coverage_Start": start_time, "Coverage_End": "",
                        "Venue": venue
                    })
                    games_found += 1
            except Exception:
                continue
        log(f"  ✅ Extracted {games_found} CFL games")
    except Exception as e:
        log(f"  ❌ CFL scrape failed: {e}")
        log(traceback.format_exc())
        raise
    finally:
        if driver:
            driver.quit()

def main():
    log("🚀 Starting web scraper")
    log_environment()
    log(f"  Target year: {TARGET_YEAR}")
    scrape_wnba()
    scrape_cfl()

    if scraped_games:
        df = pd.DataFrame(scraped_games)
        df = df.sort_values(by=['Date', 'Coverage_Start'])
        df.to_csv("games_schedule.csv", index=False)
        log(f"✅ SUCCESS: wrote {len(scraped_games)} games to games_schedule.csv")
        return

    log("❌ FAILED: no games scraped — games_schedule.csv was not created")
    sys.exit(1)

if __name__ == "__main__":
    main()


