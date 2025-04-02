# Automated-Scrapers
Automated Scrapers
You also need the following packages:

selenium
requests
beautifulsoup4
pandas
webdriver_manager
html5lib

Script Structure:        
kalibrate_report_generation(driver)
Interacts with dropdown menus on the Kalibrate site, generates several Excel reports.

carmeuse_surcharge_extraction()
Makes an HTTP request to the Carmeuse site, parses the HTML for current month surcharge info.

cn_fuel_surcharge(driver)
Clicks elements to expand and reveal a data table.
Saves the table as a CSV.

bank_of_canada_fx_download(driver)
Navigates to the Bank of Canada’s exchange rates page.
Downloads CSV of daily exchange rates, extracts the last row, and saves date + FXUSDCAD rate in a new CSV.

nrcan_wholesale_excel_download(driver)
Navigates to an NRCan page that triggers an Excel download.
Waits for a short period to allow the file to finish downloading.
