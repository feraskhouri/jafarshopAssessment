import os
import time
import re
import logging
import pandas as pd
from urllib.parse import quote

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import (
    TimeoutException,
    NoSuchElementException,
    StaleElementReferenceException,
    WebDriverException,
)

# ── CONFIGURATION ────────────────────────────────────────────────────────────

# Disable ChromeDriver/Chrome version check hack
os.environ["CHROMEDRIVER_DISABLE_BUILD_CHECK"] = "1"

# Ensure your chromedriver.exe is on the PATH
os.environ["PATH"] += r";C:\SeleniumDrivers\chromedriver-win64\chromedriver-win64"

# Logging setup
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)8s | %(message)s",
    handlers=[logging.StreamHandler()]
)

# Mapping units to conversion functions (value → grams)
UNIT_MAP = {
    'kg':     lambda v: v * 1000,
    'g':      lambda v: v,
    'lb':     lambda v: v * 453.592,
    'lbs':    lambda v: v * 453.592,
    'pound':  lambda v: v * 453.592,
    'pounds': lambda v: v * 453.592,
    'oz':     lambda v: v * 28.3495,
    'ounce':  lambda v: v * 28.3495,
    'ounces': lambda v: v * 28.3495,
}

# ── INITIALIZE BROWSER ────────────────────────────────────────────────────────

chrome_opts = webdriver.ChromeOptions()
chrome_opts.add_argument("--disable-gpu")
chrome_opts.add_argument("--no-sandbox")
chrome_opts.add_argument("--disable-dev-shm-usage")
chrome_opts.add_argument("--lang=en-US")
chrome_opts.add_argument(
    "user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
)
# chrome_opts.add_argument("--headless")  # enable for headless mode

driver = webdriver.Chrome(options=chrome_opts)
wait   = WebDriverWait(driver, 10)

# ── WEIGHT EXTRACTION FUNCTION ────────────────────────────────────────────────

def extract_weight(driver):
    """
    Attempts to extract product weight from Amazon page using structured sections.
    Returns (grams:int, source_label:str) or (None, None) if not found.
    """
    candidates = []

    # 1) Table-based details
    try:
        table = driver.find_element(By.ID, "productDetails_detailBullets_sections1")
        for row in table.find_elements(By.TAG_NAME, "tr"):
            key = row.find_element(By.TAG_NAME, "th").text.strip()
            val = row.find_element(By.TAG_NAME, "td").text.strip()
            if "weight" in key.lower():
                candidates.append((key, val))
    except NoSuchElementException:
        pass

    # 2) Bullet-list details
    try:
        ul = driver.find_element(By.ID, "detailBullets_feature_div")
        for li in ul.find_elements(By.TAG_NAME, "li"):
            parts = li.text.split(":", 1)
            if len(parts) == 2:
                key, val = parts[0].strip(), parts[1].strip()
                if "weight" in key.lower():
                    candidates.append((key, val))
    except NoSuchElementException:
        pass

    # 3) Tighter regex on each structured candidate
    pattern = re.compile(r"(\d+(?:\.\d+)?)\s*(kg|g|lbs?|pounds?|oz|ounces?)", re.I)
    for key, val in candidates:
        m = pattern.search(val)
        if m:
            num  = float(m.group(1))
            unit = m.group(2).lower().rstrip(".")
            grams = int(UNIT_MAP[unit](num))
            logging.info(f"    🔍 Found '{key}' → '{val}' → {grams} g")
            return grams, key

    # 4) No structured match
    if candidates:
        logging.warning(f"    ⚠️ Candidates found but no regex match: {candidates}")
    else:
        logging.warning("    ⚠️ No weight candidates in structured sections")
    return None, None

# ── MAIN WORKFLOW ─────────────────────────────────────────────────────────────

def main():
    # Load Excel and isolate rows needing weights
    df = pd.read_excel("product_details.xlsx")
    df = df.fillna("")
    df["Weight"] = pd.to_numeric(df["Weight"], errors="coerce").fillna(0)
    to_fill = df[df["Weight"] == 0].copy()

    # Track failures for manual review
    no_match_log = []

    for idx, row in to_fill.iterrows():
        product = row["Product Name"]
        query   = f"{product} amazon"
        logging.info(f"🔍 Searching: {query}")

        try:
            # DuckDuckGo search
            ddg_url = f"https://duckduckgo.com/?q={quote(query)}&t=h_&ia=web"
            driver.get(ddg_url)
            links = wait.until(EC.presence_of_all_elements_located((
                By.CSS_SELECTOR,
                "a.result__a, a[data-testid='result-title-a']"
            )))

            amazon_href = next(
                (a.get_attribute("href") for a in links if "amazon.com" in (a.get_attribute("href") or "").lower()),
                None
            )
            if not amazon_href:
                logging.error("❌ No Amazon link in search results.")
                no_match_log.append((product, "no_amazon_link"))
                continue

            # Navigate to Amazon
            logging.info(f"🌐 Navigate to Amazon: {amazon_href}")
            driver.get(amazon_href)
            wait.until(EC.url_contains("amazon.com"))
            time.sleep(2)

            # Extract weight
            grams, source = extract_weight(driver)

            # Fallback: full page regex
            if grams is None:
                page = driver.page_source.lower()
                fallback = re.search(
                    r"(?:item weight|shipping weight|product weight)[^\n]*?(\d+(?:\.\d*)?)\s*(kg|g|lbs?|pounds?|oz|ounces?)",
                    page
                )
                if fallback:
                    num  = float(fallback.group(1))
                    unit = fallback.group(2).lower().rstrip(".")
                    grams = int(UNIT_MAP[unit](num))
                    source = "full_page_regex"
                    logging.info(f"    🔄 Fallback regex matched → {grams} g")

            # Record result
            if grams is not None:
                df.at[idx, "Weight"] = grams
                df.at[idx, "Detection Method"] = f"ddg→amazon({source})"
                logging.info(f"✅ Parsed weight: {grams} g")
            else:
                logging.error(f"❌ Could not extract weight for '{product}'")
                no_match_log.append((product, driver.current_url))

        except (TimeoutException, NoSuchElementException, StaleElementReferenceException) as e:
            logging.error(f"❌ Timeout/Element error for '{product}': {e}")
            no_match_log.append((product, driver.current_url))
        except WebDriverException as e:
            logging.error(f"❌ WebDriver error: {e}")
            break
        except Exception as e:
            logging.error(f"❌ Unexpected error for '{product}': {e}")
            no_match_log.append((product, driver.current_url))

    # Save outputs
    df.to_excel("final_weight_ddg_amazon.xlsx", index=False)
    if no_match_log:
        pd.DataFrame(no_match_log, columns=["Product", "Note_or_URL"]) \
          .to_excel("no_match_log.xlsx", index=False)
    logging.info("✅ Done! Outputs saved.")

if __name__ == "__main__":
    try:
        main()
    finally:
        driver.quit()
