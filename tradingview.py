
import requests
from bs4 import BeautifulSoup

import bootstrap
from table import Page, SuppressLineBlock


class TradingView:
    def __init__(self, domain=''):
        URL = "https://www.tradingview.com/markets/stocks-malaysia/sectorandindustry-industry"
        if domain == '':
            domain = "electrical-products"
        self.path = "{}/{}".format(URL, domain)

    def fetch(self):
        result = []
        page = requests.get(self.path)
        soup = BeautifulSoup(page.text, 'html.parser')
        with open('console.log', "w") as bootstrap.Out:
            # out = bootstrap.Out

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

