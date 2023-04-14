from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import requests
import time
from bs4 import BeautifulSoup
import lxml
from typing import List

import bootstrap
from table import Page

global driver

# Hierarchical tree structure based on bourse, regional geo in a tree such as apac/bursa/.
# Websites to fetch URL link vary by sector and industry: -
#     https://investingmalaysia.com/category/
#     https://www.malaysiastock.biz/Listed-Companies.aspx


class TradingView:
    def __init__(self):
        # Configure the region
        region = 'usa'
        # region = 'malaysia'
        URL = f"https://www.tradingview.com/markets/stocks-{region}/sectorandindustry-"

        # combination of sector with specific domain companies
        # or simply industry with specific such as electrical-products
        sector_industry = 'sector'

        # domain = 'electrical-products'
        # domain = 'semiconductors'
        domain = 'technology-services/companies'
        self.path = f"{URL}{sector_industry}/{domain}".format(
            URL=URL, sectorandindustry=sector_industry, domain=domain)
        print("Path", self.path)

        self.load_button_text = 'loadButton-SFwfC2e0'

    def fetch(self):
        result = []
        page = requests.get(self.path)
        soup = BeautifulSoup(page.text, 'html.parser')

        # https://stackoverflow.com/questions/5041008/how-to-find-elements-by-class
        load_butt = soup.find_all("button", {"class": "{}".format(self.load_button_text)})
        if load_butt is not None:
            # cause garbage to recycle BeautifulSoup to new context
            soup = None
            result = self.fetch_chrome()
        else:
            tab = soup.find('table')
            for x in tab:
                # print(x.text)
                pager = Page(x)
                pager.parse_top()
                if len(pager.data[0]) > 0:
                    for a in pager.data:
                        result.append(a[0].lower())
        print()
        return result

    def fetch_chrome(self):
        global driver

        # bootstrap functions
        driver = webdriver.Chrome()
        session_id = driver.session_id
        executor_url = driver.command_executor._url
        driver.get(self.path)

        print("Session id:", session_id)
        print("Exec URL:", executor_url)
        print("Current cookies:", driver.get_cookies())

        while True:
            try:
                elem = driver.find_element(By.XPATH, "//button[@class='{}']".format(self.load_button_text))
                if elem is not None:
                    elem.click()
                    time.sleep(.5)
            except NoSuchElementException:
                break

        result = []
        page = driver.page_source
        soup = BeautifulSoup(page, 'lxml')

        for tab in soup.find_all('table'):
            rows = tab.find_all('tr')
            for tr in rows:
                td = tr.find_all('td')  # type: List[BeautifulSoup]
                if len(td) > 0:
                    # length of contents is asserted based on index 2 which is defined as the first class
                    class_index = 2
                    assert len(td[0].contents[0].contents) >= class_index
                    result.append(td[0].contents[0].contents[class_index].text.lower())
                    # row = [i.text for i in td]
        return result


