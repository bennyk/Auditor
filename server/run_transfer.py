from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook, worksheet
import pyperclip
import re
import time
from tkinter import *

from table import Page, fix_currency, fix_percent

global driver


class WaitCapture:
    def __init__(self):
        self.ticker = None

    def run(self):
        print("Waiting to prompt ticker")
        # https://stackoverflow.com/questions/110923/how-do-i-close-a-tkinter-window
        looper = Tk()
        Button(looper, text="Run auto", command=looper.destroy).pack()
        Button(looper, text="Exit", command=exit).pack()
        looper.mainloop()
        # elem = driver.find_element(By.ID, "input-168")
        elem = driver.find_element(By.CLASS_NAME, "v-chip__content")
        print("Found ticker", elem.text)
        self.ticker = elem.text.lower()


class MainTable:
    def __init__(self, ticker):
        self.ticker = ticker

    def open(self, title, offset):
        print("waiting to click on {title} button".format(title=title))
        time.sleep(5)

        driver.find_element(By.PARTIAL_LINK_TEXT, "{title}".format(title=title)).click()
        time.sleep(8)

        print("sliding year range to max years available")
        elem = driver.find_element(By.CLASS_NAME, "v-slider__thumb")
        print("found slider thumb")
        move = ActionChains(driver)
        # https://stackoverflow.com/questions/40485157/how-to-move-range-input-using-selenium-in-python
        # TODO -200 assuming maximized window
        # TODO range slider shorter than normal. FB/META ticker was created in 2009
        # <input value="11" id="input-min-1184" disabled="disabled" readonly="readonly" tabindex="-1">
        move.click_and_hold(elem).move_by_offset(offset, 0).release().perform()
        time.sleep(5)

    def run(self):
        clip = Clipboard()
        self.open('Financials', offset=-200)
        clip.run(selection=["Income Statement", "Balance Sheet", "Cash Flow Statement"])

        # additional offset to slide longer range span
        # TODO Values
        self.open('Valuation', offset=-880)
        clip.run(selection="Values")

        clip.save(self.ticker)
        print("Saved to {}".format(self.ticker))


class Clipboard:
    def __init__(self):
        self.wb = Workbook()

    def run(self, selection=None):
        if type(selection) is list:
            for t in selection:
                self.select(t)
                self.copy_table(t)
                self.paste(t)
        elif type(selection) is str:
            # skipping selection dialog
            t = selection
            self.copy_table(t)
            self.paste(t)

    def select(self, title):
        # https://stackoverflow.com/questions/21713280/find-div-element-by-multiple-class-names
        txt = "//*[contains(text(), '{title}')]".format(title=title)
        driver.find_element(By.XPATH, txt).click()
        # driver.find_element(By.XPATH, "//*[text()='Income Statement']").click()
        # driver.find_element(By.XPATH, "//*[text()='Balance Sheet']").click()
        # driver.find_element(By.XPATH, "//*[contains(text(), 'Cash Flow Statement')]").click()

    def copy_table(self, title):
        print("clicking 'Copy Table to Clipboard' on '{title}' table".format(title=title))
        # elem = driver.find_element(By.XPATH, "//button[contains(@class, 'v-btn v-btn--icon v-btn--round theme--light v-size--default primaryAction--text')]")
        elem = driver.find_element(By.XPATH, "//button[@class='v-btn v-btn--icon v-btn--round theme--light v-size--default primaryAction--text']")
        ActionChains(driver).click(elem).perform()
        print("copied table", elem.text)
        time.sleep(3)

    def paste(self, title):
        # https://stackoverflow.com/questions/62527396/real-time-copying-and-pasting-to-excel
        # https://python-forum.io/thread-26979.html
        clipped = pyperclip.paste()
        clipped = clipped.split('\r\n')
        clipped = [item.split('\t') for item in clipped]

        # TODO remove the previous active?
        # ws = wb.active

        # worksheet.worksheet.Worksheet
        # type: worksheet
        first_word = title.split()[0]
        ws = self.wb.create_sheet(first_word)
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
                # https://stackoverflow.com/questions/12387212/openpyxl-setting-number-format
                if number_flag:
                    cell.number_format = '0,00'
                elif per_flag:
                    cell.number_format = '0.00%'

    def save(self, ticker):
        self.wb.save('{}.xlsx'.format(ticker))


def run_main():
    global driver

    # bootstrap functions
    driver = webdriver.Chrome()
    session_id = driver.session_id
    executor_url = driver.command_executor._url
    # driver.get("https://app.tikr.com/stock/estimates?ref=iwd7tf")
    driver.get("https://app.tikr.com/login")

    print("Session id:", session_id)
    print("Exec URL:", executor_url)
    print("Current cookies:", driver.get_cookies())

    driver.maximize_window()

    elem = driver.find_element(By.ID, "input-12")
    with open('meow.txt') as f:
        if elem is not None:
            # elem.clear()
            elem.send_keys(f.readline())
            elem.send_keys(Keys.RETURN)

        elem = driver.find_element(By.ID, "input-15")
        if elem is not None:
            elem.send_keys(f.readline())
            elem.send_keys(Keys.RETURN)

    # https://www.selenium.dev/documentation/webdriver/waits/
    print("wait until username")
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "input-12"))
    )
    print("got it")

    def repeat():
        # search the text manually on Search entry manually
        capture = WaitCapture()
        capture.run()

        main = MainTable(capture.ticker)
        # run the pending chain of actions
        main.run()

    # loop forever until stop interruption
    while True:
        repeat()
        # time.sleep(3600)


if __name__ == '__main__':
    run_main()

# Footnotes
# - Selenium manpage.
#   https://selenium-python.readthedocs.io/installation.html
# - If iframe can't be solved hmmm...
#   https://www.guru99.com/handling-iframes-selenium.html
# - UI testing
#   https://success.outsystems.com/documentation/how_to_guides/devops/how_to_do_ui_testing_with_selenium/#using-selenium-for-testing-the-ui-of-your-apps
# - Requests: https://www.geeksforgeeks.org/get-post-requests-using-python/?ref=lbp
