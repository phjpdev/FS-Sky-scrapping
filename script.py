# pip install beautifulsoup4 openpyxl fake_useragent selenium
# Race Meetings.xlsx

import re
import time
import os
from pathlib import Path
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from openpyxl import load_workbook
from webdriver_manager.chrome import ChromeDriverManager

ua = UserAgent()
USER_AGENT = ua.random
ChromeDriverPath = "C:/chromedriver/chromedriver.exe"

BASE_URL = 'https://www.tab.com.au'
FILE_NAME = 'Race Meetings.xlsm'
target_column = 23
ALLOWED_MEETINGS = ['(VIC)', '(NSW)', '(QLD)', '(SA)', '(WA)', '(NT)', '(TAS)', '(ACT)', '(NZ)', '(NZL)']
FS = {}
SR = {}

def _create_chrome_service():
    """
    Prefer an explicitly configured chromedriver path when present.
    Otherwise fall back to webdriver_manager (auto-download).
    If that fails (e.g., offline/proxy), let Selenium Manager try.
    """
    configured = os.environ.get("CHROMEDRIVER_PATH") or ChromeDriverPath
    if configured:
        p = Path(configured)
        if p.is_file():
            return Service(str(p))

    try:
        return Service(ChromeDriverManager().install())
    except Exception:
        return None

def setup_driver():
    options = Options()
    options.headless = True
    options.add_argument("--disable-images")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-popup-blocking")
    options.add_argument(f"--user-agent={USER_AGENT}")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--no-first-run")
    options.add_argument("--disable-site-isolation-trials")

    service = _create_chrome_service()
    if service is None:
        driver = webdriver.Chrome(options=options)
    else:
        driver = webdriver.Chrome(service=service, options=options)
    driver.set_page_load_timeout(800)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    
    return driver

def find_all_races(html):
    soup = BeautifulSoup(html, 'html.parser')
    meetings = soup.find_all('div', {'data-testid': 'meeting', 'class': '_1e6ktkp'})
    race_links = soup.find_all('a', {'data-testid': 'race'})
    meetings_pre = [meeting.text for meeting in meetings]
    print(meetings_pre)
    meetings_names = []
    for meeting in meetings_pre:
        for allow in ALLOWED_MEETINGS:
            if allow.lower() in meeting.lower():
                meetings_names.append(meeting.split('(')[0].strip().lower())
    rounds_links = [link['href'] for link in race_links]

    print(meetings_names)

    return meetings_names, rounds_links

def extract_sky_rating(driver, url, meetings_names):
    global SR
    meeting_name = url.split('/')[3]

    if meeting_name.lower().replace('-', ' ') in meetings_names:
        SR.setdefault(meeting_name, {})
        try:
            driver.get(BASE_URL + url)
        except:
            driver.execute_script("window.stop()")

        time.sleep(5)
        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")

        # Each horse row
        rows = soup.select("div.row")  # Debug print

        for row in rows:
            try:
                horse_el = row.select_one("div.runner-name")
                if not horse_el:
                    continue

                horse_name = horse_el.get_text(strip=True).split("(")[0].strip()
                # Sky Rating is inside a <div> with a numeric value
                rating_el = row.select_one("div.runner-rating-cell span")
                if rating_el:
                    sky_rating = rating_el.get_text(strip=True)

                    if sky_rating.isdigit():
                        SR[meeting_name][horse_name] = sky_rating
                        print("Sky Ratinggggggggggggggggg:", meeting_name, horse_name, sky_rating)

            except Exception as e:
                print(f"Error extracting horse data: {e}")  # Log any errors
                continue

def extract_FS(driver, url, meetings_names):
    global FS
    meeting_name = url.split('/')[3]
    if meeting_name.lower().replace('-', ' ') in meetings_names:
        try:
            driver.get(BASE_URL + url)
        except:
            driver.execute_script("window.stop()")
        try:
            if FS[meeting_name]:
                pass
        except:
            FS[meeting_name] = {}
        try:
            button = WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[.//span[text()='Show All Form']]")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", button)
            button.click()
        except:
            pass
        while True:
            time.sleep(1)
            html = driver.page_source
            soup = BeautifulSoup(html, 'html.parser')
            FS_element = soup.find_all('p', {'class': 'comment-paragraph'})
            horse_name_divs = soup.find_all('div', {'class': 'row active'})
            if FS_element.__len__() > 0:
                print(FS_element.__len__(), horse_name_divs.__len__())
                for i in range(FS_element.__len__()):
                    horse_name = BeautifulSoup(str(horse_name_divs[i]), 'html.parser').find('div', {'class': 'runner-name'}).text.split('(')[0].strip()
                    FS[meeting_name][horse_name] = re.search(r"\(([-+]?\d*\.?\d+)\)", FS_element[i].text).group(1)
                print(FS)
                break


def get_meetings(driver, url):
    try:
        driver.get(url, )
    except:
        driver.execute_script("window.stop()")
    try:
        driver.get(url + 'R', )
    except:
        driver.execute_script("window.stop()")

    # WAIT for meeting cards
    try:
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div[data-testid='meeting']"))
        )
    except:
        print(" ERROR: TAB meetings did not load.")
        html = driver.page_source
        # print(html)
        return

    html = driver.page_source
    meetings_names, rounds_links = find_all_races(html=html)


    for i in range(rounds_links.__len__()):
        extract_FS(driver, rounds_links[i], meetings_names)
        extract_sky_rating(driver, rounds_links[i], meetings_names)


def merge_excel(excel_file, FS):
    print("\n==============================")
    print("🔍 DEBUG: Starting merge_excel")
    print("==============================")

    workbook = load_workbook(filename=excel_file, keep_vba=True)

    def normalize(name):
        return name.strip().lower().replace("-", " ")

    # Normalize all sheet names
    normalized_sheet_map = {normalize(name): name for name in workbook.sheetnames}

    print("\n📄 Sheets in workbook:")
    for k, v in normalized_sheet_map.items():
        print(f"  '{k}'  →  '{v}'")

    print("\n📌 FS meetings loaded:", list(FS.keys()))
    print("📌 SR meetings loaded:", list(SR.keys()))

    # --- PROCESS FS (TAB FS) ---
    print("\n==============================")
    print("🔸 PROCESSING TAB FS (Col W)")
    print("==============================")

    for raw_sheet_name, horses in FS.items():
        norm_name = normalize(raw_sheet_name)
        actual_sheet_name = normalized_sheet_map.get(norm_name)

        print(f"\n➡ Meeting FS: '{raw_sheet_name}' normalized to '{norm_name}'")

        if not actual_sheet_name:
            print(f"❌ No matching sheet found for FS meeting: {raw_sheet_name}")
            continue

        print(f"✔ Matched FS sheet: {actual_sheet_name}")
        print(f"🐎 Horses in FS for this meeting: {list(horses.keys())}")

        sheet = workbook[actual_sheet_name]

        for row in sheet.iter_rows(min_row=1):
            for cell in row:
                horse_name = str(cell.value).strip() if cell.value else ""
                if horse_name in horses:
                    fs_value = horses[horse_name]
                    sheet.cell(row=cell.row, column=23, value=fs_value)

                    print(f"   ➕ FS Saved | Row {cell.row} | Horse: '{horse_name}' | Value: {fs_value}")


    # --- PROCESS SKY RATING ---
    print("\n==============================")
    print("🔸 PROCESSING SKY RATING (Col X)")
    print("==============================")

    for raw_sheet_name, horses in SR.items():
        norm_name = normalize(raw_sheet_name)
        actual_sheet_name = normalized_sheet_map.get(norm_name)

        print(f"\n➡ Meeting SR: '{raw_sheet_name}' normalized to '{norm_name}'")

        if not actual_sheet_name:
            print(f"❌ No matching sheet found for SR meeting: {raw_sheet_name}")
            continue

        print(f"✔ Matched SR sheet: {actual_sheet_name}")
        print(f"🌟 Horses in SR for this meeting: {list(horses.keys())}")

        sheet = workbook[actual_sheet_name]

        for row in sheet.iter_rows(min_row=1):
            for cell in row:
                horse_name = str(cell.value).strip() if cell.value else ""
                if horse_name in horses:
                    sky_value = horses[horse_name]
                    sheet.cell(row=cell.row, column=24, value=sky_value)

                    print(f"   ⭐ Sky Saved | Row {cell.row} | Horse: '{horse_name}' | Value: {sky_value}")

    workbook.save(excel_file)
    print("\n==============================")
    print("🎉 Excel updated successfully")
    print("==============================\n")



def main():
    
    driver = setup_driver()
    get_meetings(driver=driver, url=BASE_URL + "/racing/meetings/today/")

    merge_excel(FILE_NAME, FS)

    driver.quit()


if __name__ == '__main__':
    main()
