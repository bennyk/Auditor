
import tradingview
from openpyxl import load_workbook

from spread import Spread
from utils import *
import numpy as np


class DCF(Spread):
    def __init__(self, *args):
        super().__init__(*args)

        self.wacc = .1
        self.tgr = .025
        self.term_dr = 1/(self.wacc-self.tgr)

    def wa_diluted_shares_out(self) -> float:
        result = self.strip(self.income.match_title('Weighted Average Diluted Shares Outstanding'))
        return result

    def fcf(self):
        # also known as Levered FCF
        earning_not_strip = self.cashflow.match_title('Free Cash Flow$', none_is_optional=True)
        if earning_not_strip is not None:
            earning = self.strip(earning_not_strip)
        else:
            cfo = self.strip(self.cashflow.match_title('Cash from Operations$'))
            opt_acq_real_assets = self.cashflow.match_title('Acquisition of Real Estate Assets$',
                                                            none_is_optional=True)
            earning = cfo
            if opt_acq_real_assets is not None:
                acq_real_assets = self.strip(opt_acq_real_assets)
                earning = list_add_list(cfo, acq_real_assets)
        return list_over_list(earning, self.wa_diluted_shares_out())

    def tv_ebitda(self, _):
        # hmm, not preferred when comparing company. The preferred is EBITDA
        # ni = self.strip(self.income.match_title('Net Income$'))
        # tax = self.strip(self.income.match_title('Income Tax Expense$'))
        # ie = self.strip(self.income.match_title('Interest Expense$'))
        # ebit = list_add_list(ni, tax)
        # ebit = list_add_list(ebit, ie)
        # ebit_per_share = list_over_list(ebit, self.wa_diluted_shares_out())

        ebitda = self.strip(self.income.match_title('EBITDA'))
        ebitda_per_share = list_over_list(ebitda[self.half_len:], self.wa_diluted_shares_out())
        return average(ebitda_per_share)*(1+self.tgr) * self.term_dr

    def tv_last_fcf(self, fcf):
        return fcf[-1]*(1+self.tgr) * self.term_dr

    def tv_avg_fcf(self, fcf):
        # tv = average(fcf)*(1+tgr) * term_dr
        return average(fcf[self.half_len:])*(1+self.tgr) * self.term_dr

    def compute_fcf(self):
        print(self.tick)
        fcf = self.fcf()
        nper = len(fcf)

        # https://www.investopedia.com/terms/d/dcf.asp
        dr = [1/(1+self.wacc) ** (i+1) for i in range(nper)]
        pv = [fcf[i] * dr[i] for i in range(nper)]

        print('nper', nper)
        print('cagr of fcf', '{:.1f}%'.format(cagr(fcf)*100))
        print('avg of fcf', '{:.1f}%'.format(average(fcf)*100))
        print('fcf', np.around(fcf, decimals=2))
        print('dr', np.around(dr, decimals=2))
        print('pv', np.around(pv, decimals=2))
        print('term_dr', '{:.2f}'.format(self.term_dr))

        # https://www.investopedia.com/terms/t/terminalvalue.asp
        tv = []
        sum_pvtv = []
        for m in [self.tv_avg_fcf, self.tv_last_fcf, self.tv_ebitda]:
            _ = m(fcf)
            tv.append(_)
            sum_pvtv.append(sum(pv)+_)

        # TODO table fix
        # https://blog.devgenius.io/how-to-easily-print-and-format-tables-in-python-18bbe2e59f5f
        print('tv\t\t\t', np.around(tv, decimals=2))
        print('sum_pvtv\t', np.around(sum_pvtv, decimals=2))
        print()

        # adbe
        # fcf = [15.6, 18.72, 21.53, 24.76, 28.47, 30.6, 32.9, 35.36, 38.02, 40.87, 708.47]


# tickers = tradingview.TradingView().fetch()
tickers = ['adbe', 'intu', 'sap', 'intc', 'qcom', 'googl', 'meta', 'txn', 'mu', 'asml', 'cdb', 'tm']
path = "spreads"
for tick in tickers:
    wb = load_workbook(path + '/' + tick + '.xlsx')
    spread = DCF(wb, tick)
    spread.compute_fcf()
