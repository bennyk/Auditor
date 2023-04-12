
import requests
from bs4 import BeautifulSoup

import bootstrap
from table import Page

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
        # sector_industry = 'sector'
        # domain = 'consumer-services/companies'
        # or
        sector_industry = 'industry'
        # domain = 'electrical-products'
        domain = 'semiconductors'
        self.path = f"{URL}{sector_industry}/{domain}".format(
            URL=URL, sectorandindustry=sector_industry, domain=domain)

    def fetch(self):
        result = []
        page = requests.get(self.path)
        soup = BeautifulSoup(page.text, 'html.parser')
        with open('console.log', "w") as bootstrap.Out:
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

