# jafarshopAssessment


### **Project Summary: Xiaomi Product Weight Collection**

#### **1. Data Extraction and Initial Processing**

* Parsed the provided HTML file to extract structured product information.
* Created an Excel file `product_detalis.xlsx` containing:

  * Product Name (EN)
  * Model Number
  * Weight
  * Detection Method

#### **2. Data Cleaning**

* Identified and removed duplicate entries (e.g., color variations).
* Saved the cleaned dataset as `product_detalis_cleaned.xlsx`.

#### **3. Web Scraping – Amazon.com via DuckDuckGo**

* Used **DuckDuckGo** to search for each product by name along with the keyword "Amazon".
* Clicked on the first Amazon product link from the search results.
* Scraped the **Item Weight** from the product’s detail section.
* Logged and merged weight data into the dataset.

#### **4. Web Scraping – mi.com (Specifications Section)**

* Developed a Selenium-based scraper to:

  * Search each product name on [mi.com/global](https://www.mi.com/global/).
  * Navigate to the **Specifications** tab of the first matching result.
  * Extract and store product weight when available.

#### **5. Web Scraping – mi.com (Support Section)**

* Retrieved weight data from **Support** pages for products not listed in the Specifications section.
* Merged the newly found weights with the existing dataset.

#### **6. Manual Entry**

* Manually entered weight data for 4 products that couldn't be found through automated methods.

