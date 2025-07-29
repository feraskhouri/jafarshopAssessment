import os
import time
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException
from concurrent.futures import ThreadPoolExecutor

# ── 1. Chrome Driver Setup ──────────────────────────────────────────────────────
chrome_options = Options()
chrome_options.add_argument('--disable-blink-features=AutomationControlled')
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)
chrome_options.add_argument('--start-maximized')
os.environ['PATH'] += r";C:\SeleniumDrivers\chromedriver-win64\chromedriver-win64"

# ── 2. Load & Prepare DataFrame ────────────────────────────────────────────────
df = pd.read_excel("final.xlsx").fillna('')
df['Weight'] = pd.to_numeric(df['Weight'], errors='coerce').fillna(0)
if 'Detection Method' not in df.columns:
    df['Detection Method'] = ''
if 'WeightUnit' not in df.columns:
    df['WeightUnit'] = ''

# We'll only scrape those with no recorded weight yet
to_scrape = df[df['Weight'] == 0].copy()

# ── 3. Precompile Regex ────────────────────────────────────────────────────────
weight_regex = re.compile(
    r'(\b\d+(?:\.\d+)?\s*(?:kg|kilograms|g|grams)\b)', 
    re.IGNORECASE
)

# ── 4. Helper: Parse weight text into (value, unit) ────────────────────────────
def parse_weight_text(text):
    txt = text.lower()
    unit = None
    if 'kg' in txt:
        unit = 'kg'
    elif 'g' in txt:
        unit = 'g'
    else:
        return None, None

    num_match = re.search(r'[\d.]+', txt)
    if not num_match:
        return None, None

    try:
        value = float(num_match.group())
    except ValueError:
        return None, None

    return value, unit

# ── 5. Core Scraping Logic per Product ─────────────────────────────────────────
def process_row(idx, row):
    name = row['Product Name']
    print(f"\n🔍 Searching support for: {name}")
    for attempt in (1, 2):
        driver = webdriver.Chrome(options=chrome_options)
        wait = WebDriverWait(driver, 15)
        try:
            driver.get("https://www.mi.com/global/")
            wait.until(EC.presence_of_element_located((By.ID, 'mi-base-search')))
            time.sleep(1)

            # Dismiss any pop-up
            try:
                popup = WebDriverWait(driver, 3).until(
                    EC.element_to_be_clickable((By.CLASS_NAME, "shortcut__item--wrapper"))
                )
                popup.click()
            except:
                pass

            # Perform search
            search_input = wait.until(EC.element_to_be_clickable((By.ID, 'mi-base-search')))
            search_input.clear()
            search_input.send_keys(f"{name} weight")
            search_input.send_keys(Keys.ENTER)
            time.sleep(2)

            # Click Support tab
            support_tab = wait.until(EC.element_to_be_clickable(
                (By.CSS_SELECTOR, 'li.search-tabs--item[data-tab-type="support"]')
            ))
            support_tab.click()
            time.sleep(2)

            # 1) Preview scrape
            previews = driver.find_elements(By.CSS_SELECTOR, 'div.support-result-item__left')
            snippet = " ".join([el.text for el in previews]).lower()
            m = weight_regex.search(snippet)
            if m:
                val, unit = parse_weight_text(m.group(1))
                if val is not None:
                    print(f"✅ Found {val} {unit} in preview")
                    driver.quit()
                    return idx, val, unit, 'support-preview'

            # 2) Full page scrape (click first result)
            for _ in range(3):
                try:
                    link = wait.until(EC.element_to_be_clickable(
                        (By.CSS_SELECTOR, 'a.support-result-item__left--link')
                    ))
                    link.click()
                    break
                except StaleElementReferenceException:
                    time.sleep(1)

            wait.until(EC.presence_of_element_located((By.TAG_NAME, 'body')))
            body_text = driver.find_element(By.TAG_NAME, 'body').text.lower()
            m = weight_regex.search(body_text)
            if m:
                val, unit = parse_weight_text(m.group(1))
                if val is not None:
                    print(f"✅ Found {val} {unit} in full page")
                    driver.quit()
                    return idx, val, unit, 'support-full'

            print("⚠️ No weight found on support pages.")
            driver.quit()
            return idx, None, None, ''

        except TimeoutException:
            print(f"⏳ Timeout on attempt {attempt} for '{name}'")
            try:
                with open(f"failed_{idx}.html", "w", encoding="utf-8") as f:
                    f.write(driver.page_source)
            except:
                pass
            driver.quit()
            time.sleep(2)

        except Exception as e:
            print(f"❌ Error for '{name}': {e}")
            driver.quit()
            return idx, None, None, ''

    # If both attempts fail
    return idx, None, None, ''

# ── 6. Execute in Parallel & Collect Results ──────────────────────────────────
results = []
with ThreadPoolExecutor(max_workers=4) as exe:
    futures = [exe.submit(process_row, i, r) for i, r in to_scrape.iterrows()]
    for fut in futures:
        res = fut.result()
        if res:
            results.append(res)

# ── 7. Write Back to DataFrame & Save ──────────────────────────────────────────
for idx, val, unit, method in results:
    if val is not None:
        df.at[idx, 'Weight'] = val
        df.at[idx, 'WeightUnit'] = unit
        df.at[idx, 'Detection Method'] = method

df.to_excel("done.xlsx", index=False)
print("\n✅ Done. Saved as 'weigh.xlsx'.")