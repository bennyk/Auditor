# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

from openpyxl import load_workbook
from reit2 import ProfManager, Spread


def main():
    path = "C:/Users/benny/iCloudDrive/Documents/company-spreads"
    prof = ProfManager()

    # tickers = ['intc', 'tsm', 'nvda', 'amd', 'txn', 'qcom', 'mu', 'csco',
    #            'meta', 'kr', 'mdt', 'tsla', 'aapl', 'msft', 'adi', 'goog', 'brk-b',
    #            '3333', '1810',
    #            'kipreit', 'icap', 'ghlsys', 'digi', 'genting', 'mieco',
    #            'igbreit', 'kobay', 'dpharma', 'timecom',
    #            'slb', 'oxy', 'xom',
    #            #'mob',
    #            ]
    # tickers = ['intc', 'nvda', 'brk-b', 'ghlsys', 'digi','revenue',  ]
    tickers = ['vs', 'sam', 'skpres', 'uchitec', 'pie', 'dufu', 'kobay',
               'wellcal', 'cbip', 'chinwel', 'boilerm', 'qes',
               # 'ataims', 'qes'

   # tickers = ['digi', 'maxis', 'axiata', 'tm', 'timecom', 'redtone', 'ock', ]
               # 'gpacket', 'xox'
               'inari', 'vitrox', 'mpi', 'd&o', 'unisem', 'frontkn', 'uwc', 'gtronic',
               'jhm', 'kesm', 'vis', 'keyasic',

               'greatec', 'penta', 'genetec', 'mi',
               'genting',
               'mfcb',]

    # tickers = ['mfcb']

    # TODO Adding TODO may need to fix AHP.
    for c in tickers:
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

