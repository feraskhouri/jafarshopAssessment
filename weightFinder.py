import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
import re

def extract_model_number(driver):
    """Extract model number using Selenium, returns ' ' if unavailable"""
    paragraphs = driver.find_elements(By.TAG_NAME, "p")
    for p in paragraphs:
        if "رقم الموديل" in p.text:
            model_text = p.text.split(":")[-1].strip()
            if "غير متوفر" in model_text or not model_text:
                return " "
            return model_text
    return " "

def extract_weight(driver):
    """Comprehensive weight extraction using Selenium"""
    full_text = driver.find_element(By.TAG_NAME, "body").text
    weight_patterns = [
        r"الوزن[\s:]*حوالي\s*([\d\.]+)\s*جرام",
        r"الوزن[\s:]*([\d\.]+)\s*جرام",
        r"وزنه[\s:]*([\d\.]+)\s*جرام",
        r"([\d\.]+)\s*جرام",
        r"([\d\.]+)\s*g\b"
    ]
    for pattern in weight_patterns:
        match = re.search(pattern, full_text, re.IGNORECASE)
        if match:
            try:
                return float(match.group(1)), "Automatic Extraction"
            except ValueError:
                continue
    return 0, "Not Detected"

# Set up Selenium WebDriver (you may need to specify the path to your driver)
driver = webdriver.Chrome()  # or webdriver.Firefox(), etc.

# Load data
df = pd.read_csv("Products weight .html")
results = []

for _, row in df.iterrows():
    # Load HTML content into Selenium
    driver.get("data:text/html;charset=utf-8," + row['Body (HTML)'])
    
    # Extract English name
    name_en = "Unknown"
    paragraphs = driver.find_elements(By.TAG_NAME, "p")
    for p in paragraphs:
        if "اسم المنتج بالإنجليزي" in p.text:
            name_en = p.text.split(":")[-1].strip()
            break
    
    # Extract model number
    model_num = extract_model_number(driver)
    
    # Extract weight and method
    weight, method = extract_weight(driver)
    
    results.append({
        "Product Name": name_en,
        "Model Number": model_num,
        "Weight": weight,
        "Detection Method": method
    })
# Close the browser
driver.quit()

# Save to Excel
output_df = pd.DataFrame(results)
output_df.to_excel("product_details.xlsx", index=False)
print(" Done! File saved with detection methods column.")