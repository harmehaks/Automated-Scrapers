import os
import csv
import time
import requests
import pandas as pd

from datetime import datetime
from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import Select

def kalibrate_report_generation(driver):
    print("\n=== Starting Kalibrate Report Generation ===")

    driver.get("https://charting.kalibrate.com/")
    try:
        wait = WebDriverWait(driver, 15)
        wait.until(EC.visibility_of_element_located((By.ID, "Product")))
    except:
        print("Could not find the Product dropdown. Aborting Kalibrate script.")
        return

    input_options_list = [
        "Diesel/Wholesale/Daily/Current year/Excel",
        "Diesel/Retail/Weekly/Current year/Excel",
        "Diesel/Retail excl tax/Weekly/Current year/Excel",
        "Taxation/Taxation Tables/Historical Taxation/Current year/Excel",
        "Diesel/Wholesale by Marketer/Daily/Current year/Excel"
    ]

    for input_string in input_options_list:
        try:
            options = input_string.split('/')
            Select(driver.find_element(By.ID, 'Product')).select_by_visible_text(options[0])
            time.sleep(1)
            Select(driver.find_element(By.ID, 'ReportType')).select_by_visible_text(options[1])
            time.sleep(1)
            Select(driver.find_element(By.ID, 'Frequency')).select_by_visible_text(options[2])
            time.sleep(1)
            current_year = datetime.now().year
            time.sleep(1)
            Select(driver.find_element(By.ID, 'Year')).select_by_visible_text(str(current_year))
            time.sleep(1)
            Select(driver.find_element(By.ID, 'ReportGenerationType')).select_by_visible_text(options[4])
            time.sleep(1)
            generate_report_button = driver.find_element(By.ID, 'GenerateReport')
            generate_report_button.click()

            time.sleep(10)

            print(f"Generated report for: {input_string}")

        except Exception as e:
            print(f"Error generating report for {input_string}: {e}")

    print("=== Kalibrate Reports Completed ===")

def carmeuse_surcharge_extraction():
    print("\n=== Starting Carmeuse Surcharge Extraction ===")

    url = "https://www.carmeuse.com/na-en/services/freight-logistics"
    current_year = str(datetime.now().year)
    current_month = datetime.now().strftime("%b")

    try:
        response = requests.get(url)
        if response.status_code != 200:
            print(f"Failed to retrieve {url}. Status code: {response.status_code}")
            return

        soup = BeautifulSoup(response.text, 'html.parser')
        tables = soup.find_all('table')
        if not tables:
            print("No <table> elements found on the page.")
            return

        for table in tables:
            title = table.find_previous(['h3', 'h2', 'p'])
            title_text = title.get_text() if title else ""

            if current_year in title_text:
                df = pd.read_html(str(table))[0]
                if current_month in df.iloc[:, 0].values:
                    current_month_data = df[df.iloc[:, 0] == current_month].iloc[0]
                    us_surcharge = current_month_data[5]
                    canada_surcharge = current_month_data[6]

                    print(f"US Surcharge: {us_surcharge}")
                    print(f"Canada Surcharge: {canada_surcharge}")

                    filename = f"price_of_surcharge_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
                    data = {
                        "Month": [current_month],
                        "US Surcharge": [us_surcharge],
                        "Canada Surcharge": [canada_surcharge]
                    }
                    df_to_save = pd.DataFrame(data)
                    df_to_save.to_csv(filename, index=False)
                    print(f"Data saved to {filename}")
                    return

        print(f"No table found with the current year ({current_year}) in the title.")
    except Exception as e:
        print(f"Error parsing Carmeuse site: {e}")


def cn_fuel_surcharge(driver):
    print("\n=== Starting CN Fuel Surcharge Extraction ===")

    url = "https://www.cn.ca/en/customer-centre/prices-tariffs-transit-times/fuel-surcharge/"
    driver.get(url)

    try:
        wait = WebDriverWait(driver, 15)

        try:
            cookie_customize_button = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button.cky-btn.cky-btn-reject"))
            )
            cookie_customize_button.click()
            print("Cookie banner dismissed.")
        except:
            print("No cookie banner or couldn't dismiss it. Continuing...")

        expand_buttons = wait.until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, ".open"))
        )
        print(f"Found {len(expand_buttons)} expandable element(s). Clicking them...")

        for button in expand_buttons:
            button.click()
            time.sleep(1)


        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table")))

        page_source = driver.page_source
        soup = BeautifulSoup(page_source, 'html.parser')

        first_table = soup.find('table')
        if not first_table:
            print("No table found on the CN page.")
            return

        rows = first_table.find_all('tr')
        if len(rows) < 4:
            print("Table doesn't have enough rows to parse as expected.")
            return

        header_row = rows[0]
        header_cols = header_row.find_all(['th', 'td'])
        if header_cols:
            table_name = header_cols[0].get_text(strip=True)
        else:
            table_name = "CN_fuel_surcharge"

        current_time = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"{table_name}_{current_time}.csv"

        header = [
            'Effective Month',
            'Average OHD Price',
            'Intra-Canada Bulk',
            'Intra-Canada Carload',
            'U.S. Carload (incl. Bulk)'
        ]

        with open(filename, mode='w', newline='', encoding='utf-8') as file:
            writer = csv.writer(file)
            writer.writerow(header)
            for row in rows[3:]:
                cols = row.find_all('td')
                col_text = [col.get_text(strip=True) for col in cols]
                writer.writerow(col_text)

        print(f"Data saved to '{filename}'")

    except Exception as e:
        print(f"Error scraping CN Fuel Surcharge: {e}")


def bank_of_canada_fx_download(driver):
    print("\n=== Starting Bank of Canada FX Download ===")

    url = "https://www.bankofcanada.ca/rates/exchange/daily-exchange-rates/"
    driver.get(url)

    try:
        wait = WebDriverWait(driver, 15)
        download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//a[contains(text(), 'CSV')]")))
        download_button.click()
        print("Clicked the CSV download link...")

        time.sleep(5)

        download_folder = os.path.join(os.path.expanduser("~"), "Downloads")
        files = os.listdir(download_folder)
        csv_files = [os.path.join(download_folder, f) for f in files if f.endswith('.csv')]
        if not csv_files:
            print("No CSV file found in the Downloads folder.")
            return

        latest_csv_file = max(csv_files, key=os.path.getctime)
        print(f"Latest CSV file: {latest_csv_file}")

        df = pd.read_csv(latest_csv_file, skiprows=39, on_bad_lines='skip')
        if "FXUSDCAD" not in df.columns or "date" not in df.columns:
            print("The downloaded CSV does not have the expected 'FXUSDCAD' or 'date' columns.")
            return

        latest_value = df["FXUSDCAD"].iloc[-1]
        latest_date = df["date"].iloc[-1]

        print(f"Latest date for FXUSDCAD: {latest_date}")
        print(f"Value of FXUSDCAD on the latest date: {latest_value}")

        result_df = pd.DataFrame({
            'date': [latest_date],
            'FXUSDCAD': [latest_value]
        })

        current_time = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        filename = f"most_recent_posted_US_dollar_price_{current_time}.csv"
        result_df.to_csv(filename, index=False)
        print(f"Data saved in: {filename}")

    except Exception as e:
        print(f"Error while reading or downloading Bank of Canada CSV: {e}")


def nrcan_wholesale_excel_download(driver):
    print("\n=== Starting NRCan Wholesale Excel Download ===")

    current_year = datetime.now().year
    url = f"https://www2.nrcan.gc.ca/eneene/sources/pripri/wholesale_bycity_e.cfm?priceYear={current_year}&productID=13&locationID=9,74&frequency=M&downloadXLS"

    try:
        driver.get(url)
        time.sleep(10)
        print("Page loaded. The Excel file should have started downloading (if triggered by the site).")

    except Exception as e:
        print(f"An error occurred: {e}")

def main():
    print("\n=== Script Started ===")
    carmeuse_surcharge_extraction()
    chrome_options = Options()
    prefs = {
        "download.default_directory": os.path.join(os.path.expanduser("~"), "Downloads"),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
    }
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    try:
        kalibrate_report_generation(driver)
        cn_fuel_surcharge(driver)
        bank_of_canada_fx_download(driver)
        nrcan_wholesale_excel_download(driver)
    finally:
        print("\nClosing the browser...")
        driver.quit()
    print("=== All tasks completed ===")

if __name__ == "__main__":
    main()