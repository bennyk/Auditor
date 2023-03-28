# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import os.path

from openpyxl import load_workbook
from tikr_terminal import ProfManager, Spread
from bcolors import bcolors, colour_print
import tradingview


def main():
    prof = ProfManager()
    tickers = tradingview.TradingView().fetch()

    path = "spreads"
    for c in tickers:
        xls_path = path+'/' + c + '.xlsx'
        if not os.path.isfile(xls_path):
            colour_print("Excel file is missing: '{}'".format(xls_path), bcolors.WARNING)
            continue

        print('Ticker {}'.format(c))
        wb = load_workbook(path+'/' + c + '.xlsx')
        pf = prof.create_folder(c)
        t = Spread(wb, c, pf)
        t.revenue()
        t.epu()
        t.owner_yield()
        # t.cfo()
        # AFFO commented diff
        # t.affo()
        # t.nav()
        # Tangible commented diff
        # t.tangible_book()
        # t.return_equity()
        t.return_invested_cap()
        t.net_debt_over_ebit()
        # t.net_debt_over_fcf()
        t.retained_earnings_ratio()
        t.market_cap_over_retained_earnings_ratio()
        t.ebit_margin()
        t.ev_over_ebit()
        # t.dividend_payout_ratio()
        t.div_yield()
        t.last_price()
        print()

    prof.profile()


if __name__ == '__main__':
    main()

