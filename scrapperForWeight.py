import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import re

# Set up ChromeDriver path
os.environ['PATH'] += r";C:\SeleniumDrivers\chromedriver-win64\chromedriver-win64"
#chrome_options.add_argument("--headless")

# Load Excel and prepare DataFrame
df = pd.read_excel("final_weight_ddg_amazon.xlsx")
df = df.fillna('')
df['Weight'] = pd.to_numeric(df['Weight'], errors='coerce').fillna(0)
filtered_df = df[df['Weight'] == 0].copy()

# Launch browser
driver = webdriver.Chrome()
wait = WebDriverWait(driver, 10)

# Loop through each product with missing weight
for index, row in filtered_df.iterrows():
    product_name = row['Product Name']
    print(f"\n🔍 Searching for: {product_name}")

    try:
        # Always go to Xiaomi homepage fresh for each search
        driver.get("https://www.mi.com/global/")
        wait.until(EC.presence_of_element_located((By.ID, 'mi-base-search')))
        time.sleep(1)

        # Handle popup shortcut if exists
        try:
            shortcut_item = WebDriverWait(driver, 3).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "shortcut__item--wrapper"))
            )
            shortcut_item.click()
        except:
            pass

        # Search for the product
        search_box = wait.until(EC.element_to_be_clickable((By.ID, 'mi-base-search')))
        search_box.clear()
        search_box.send_keys(product_name)
        search_box.send_keys(Keys.ENTER)

        time.sleep(3)

        # Click first result
        first_item = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, ".product-result-item")))
        first_item.click()

        # Go to Specs tab
        specs_link = wait.until(EC.element_to_be_clickable((By.ID, "nav-specs")))
        specs_link.click()

        time.sleep(3)

        # Look for weight-related spans
        weight_elements = driver.find_elements(By.CSS_SELECTOR, "span.xm-text, span.xm-text.f-light")
        found = False

        for el in weight_elements:
            text = el.text.strip().lower()
            if any(keyword in text for keyword in ['weight', 'net weight','weight:','g approx','Product net weight']) or re.search(r'\b\d+(\.\d+)?\s*(g|kg)\b', text, re.IGNORECASE):
                print(f" Found potential weight text: {text}")

                if 'kg' in text:
                    try:
                        kg_val = float(''.join(c if c.isdigit() or c == '.' else '' for c in text))
                        weight_val = int(kg_val * 1000)
                        df.loc[index, 'Weight'] = weight_val
                        df.loc[index, 'Detection Method'] = 'scraper'
                        print(f" Parsed weight: {weight_val}g (from kg)")
                        found = True
                        break
                    except:
                        continue
                elif 'g' in text:
                    try:
                        g_val = int(''.join(filter(str.isdigit, text)))
                        df.loc[index, 'Weight'] = g_val
                        df.loc[index, 'Detection Method'] = 'scraper'
                        print(f" Parsed weight: {g_val}g")
                        found = True
                        break
                    except:
                        continue

        if not found:
            print(" Weight not found or format unrecognized.")

    except Exception as e:
        print(f" Failed for '{product_name}': {e}")
        continue

# Save the updated file
df.to_excel("final.xlsx", index=False)
print("\n✅ Done! Updated Excel saved as 'product_detailsS_updated.xlsx'.")
# Close the browser
driver.quit()  