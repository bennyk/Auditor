import time

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook, worksheet
import pyperclip
import re

from table import Page, fix_currency, fix_percent

driver = webdriver.Chrome()
session_id = driver.session_id
executor_url = driver.command_executor._url
# driver.get("https://app.tikr.com/stock/estimates?ref=iwd7tf")
driver.get("https://app.tikr.com/login")

print("Session id:", session_id)
print("Exec URL:", executor_url)
print("Current cookies:", driver.get_cookies())

elem = driver.find_element(By.ID, "input-12")
if elem is not None:
    # elem.clear()
    elem.send_keys("benny_khoo_99@yahoo.com")
    elem.send_keys(Keys.RETURN)

elem = driver.find_element(By.ID, "input-15")
if elem is not None:
    elem.send_keys("cismop-xUqkeh-sapdo5")
    elem.send_keys(Keys.RETURN)

print("wait until username")
element = WebDriverWait(driver, 10).until(
    EC.presence_of_element_located((By.ID, "input-12"))
)
print("got it")

# TODO search the text manually on Search entry manually
time.sleep(5)

print("waiting to click on Financials button")
time.sleep(5)

driver.find_element(By.PARTIAL_LINK_TEXT, "Financials").click()
# WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.PARTIAL_LINK_TEXT, "Financials")))
time.sleep(5)

print("sliding year range to max years available")
elem = driver.find_element(By.CLASS_NAME, "v-slider__thumb")
move = ActionChains(driver)
# TODO -200 assuming maximized window
move.click_and_hold(elem).move_by_offset(-200, 0).release().perform()
time.sleep(5)

print("clicking 'Copy Table to Clipboard'")
# elem = driver.find_element(By.XPATH, "//button[contains(@class, 'v-btn v-btn--icon v-btn--round theme--light v-size--default primaryAction--text')]")
elem = driver.find_element(By.XPATH, "//button[@class='v-btn v-btn--icon v-btn--round theme--light v-size--default primaryAction--text']")
ActionChains(driver).click(elem).perform()
print("copied table", elem.text)
time.sleep(5)

clipped = pyperclip.paste()
clipped = clipped.split('\r\n')
clipped = [item.split('\t') for item in clipped]
wb = Workbook()

# TODO remove the previous active?
# ws = wb.active

# worksheet.worksheet.Worksheet
# type: worksheet
ws = wb.create_sheet("Income")
# created sheet may intro the title automatically
# ws.title = "Income Statement"

for row, row_data in enumerate(clipped, start=1):
    for col, cell_data in enumerate(row_data, start=1):
        # print(cell_data)
        number_flag = False
        per_flag = False
        if re.match(Page.re_numerical, cell_data):
            try:
                if re.match(Page.re_percent, cell_data):
                    cell_data = fix_percent(cell_data)
                    per_flag = True
                else:
                    # assuming it is a number until it caught exception
                    cell_data = fix_currency(cell_data)
                    number_flag = True
            except ValueError:
                pass
        cell = ws.cell(row=row, column=col, value=cell_data)
        if number_flag:
            cell.number_format = '0,00'
        elif per_flag:
            cell.number_format = '0.00%'
wb.save('spam.xlsx')

time.sleep(3600)
