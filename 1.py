import os
import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# -------- SETTINGS ----------
HEADLESS_MODE = False
TOP_COINS = 10
USD_TO_INR = 90  # You can update conversion rate if needed

# Always save to Desktop
OUTPUT_FILE = r"C:\Users\jaisa\OneDrive\Desktop\crypto_prices_inr.xlsx"
# ----------------------------


def setup_driver():
    options = Options()
    if HEADLESS_MODE:
        options.add_argument("--headless")
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install()),
        options=options
    )
    return driver


def scrape_data():
    driver = setup_driver()
    driver.get("https://coinmarketcap.com/")

    # Wait until table loads properly
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "//table/tbody/tr"))
    )

    rows = driver.find_elements(By.XPATH, "//table/tbody/tr")[:TOP_COINS]

    data = []

    for row in rows:
        try:
            name = row.find_element(By.XPATH, ".//p[contains(@class,'coin-item-symbol')]").text
            price_text = row.find_element(By.XPATH, ".//td[4]").text
            change_24h = row.find_element(By.XPATH, ".//td[5]").text
            market_cap = row.find_element(By.XPATH, ".//td[8]").text

            # Clean price (remove symbols and commas)
            numeric_price = float(
                price_text.replace("$", "")
                          .replace("₹", "")
                          .replace(",", "")
            )

            # Convert USD to INR if needed
            if "$" in price_text:
                numeric_price *= USD_TO_INR

            price_inr = f"₹{numeric_price:,.2f}"

            data.append([
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                name,
                price_inr,
                change_24h,
                market_cap
            ])

        except Exception as e:
            print("Row skipped:", e)

    driver.quit()

    df = pd.DataFrame(data, columns=[
        "Timestamp",
        "Coin",
        "Price (INR)",
        "24h Change",
        "Market Cap"
    ])

    return df


def save_to_excel(df):
    df.to_excel(OUTPUT_FILE, index=False)
    print("\n✅ Data successfully saved to:")
    print(os.path.abspath(OUTPUT_FILE))


if __name__ == "__main__":
    print("📊 Fetching live crypto prices...")

    df = scrape_data()

    if not df.empty:
        print(df)
        save_to_excel(df)
    else:
        print("⚠ No data extracted.")