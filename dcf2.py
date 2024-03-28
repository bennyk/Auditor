import re
import pandas as pd

from spread import Spread
from utils import *
from bcolors import colour_print, bcolors

from openpyxl import Workbook, load_workbook, worksheet
from openpyxl.styles import Font, Alignment
from collections import OrderedDict
from typing import List

class ExcelOut:
    start_col = 2
    row_margin = 1

    def __init__(self, tick: str, entries, styles=None, headers=None):
        cls = self.__class__
        self.od = entries
        self.headers = headers
        self.tick = tick
        self.styles = styles
        self.ft = Font(name='Calibri', size=11)
        self.wb = Workbook()

        # type: worksheet.Worksheet
        self.sheet = self.wb.active
        self.sheet.title = 'sheet 1'

        self.cell = self.sheet.cell(row=1, column=cls.start_col)
        self.start_row_index = cls.row_margin+1
        self.end_row_index = len(self.tick) + self.start_row_index+1

        self.j = cls.row_margin + 1
        self.i = 1

        self.init_sheet()

    def init_sheet(self):
        sheet = self.sheet
        self.j = 1
        i = 2
        for val in self.headers:
            sheet.column_dimensions[colnum_string(self.j)].width = 32
            cell = sheet.cell(row=i, column=self.j)
            cell.alignment = Alignment(wrapText=True)
            if re.match(r'empty', val):
                pass
            else:
                cell.value = val
            i += 1

        cell = sheet.cell(row=1, column=1)
        cell.value = self.tick.upper()
        for i in range(2, 12):
            cell = sheet.cell(row=1, column=i)
            cell.value = i-1
            cell.alignment = Alignment(horizontal='center')
        cell = sheet.cell(row=1, column=i+1)
        cell.value = 'Terminal year'

        self.j += 1

    def start(self):
        cls = self.__class__
        self.j = cls.row_margin + 1
        for i, key in enumerate(self.od):
            # Table of mainly profile and last_price data
            self.i = 2
            if type(self.od[key]) is list:
                for val in self.od[key]:
                    if key != '':
                        self.make_cell(val, self.styles[i])
            else:
                assert type(self.od[key]) is float or type(self.od[key]) is int
                val = self.od[key]
                self.make_cell(val, self.styles[i])

            self.j += 1
        self.wb.save('out.xlsx')

    def make_cell(self, val, style):
        cell = self.sheet.cell(row=self.j, column=self.i)
        self.sheet.column_dimensions[colnum_string(self.i)].width = 11
        cell.alignment = Alignment(wrapText=True)
        cell.value = val
        if style == 'Comma':
            if val != 0:
                cell.style = style
                cell.number_format = '#,0.00'
        elif style == 'Percent':
            cell.style = style
            cell.number_format = '0.00%'
        elif style == 'Ratio2':
            cell.number_format = '0.00'
        elif style == 'Ratio':
            # cell.style = style
            cell.number_format = '0.0000'
        else:
            assert False
        cell.font = self.ft
        self.i += 1


class DCF(Spread):
    def __init__(self, tick, path):
        self.wb = load_workbook(path + '/' + tick + '.xlsx')
        super().__init__(self.wb, tick)

        # Revenues, Operating Income, Interest Expense, ...

        self.forward_sales = self.trim_estimates('Revenue', nlead=8)
        self.forward_ebit = self.trim_estimates('EBIT$')

        # TODO Interest expense and Equity?
        # self.forward_ie = self.strip(self.estimates.match_title('Interest Expense'))
        # self.equity = self.strip(self.balance.match_title('Total Equity'))
        self.debt = self.strip(self.balance.match_title('Total Debt'))

        cash_not_strip = self.balance.match_title('Total Cash', none_is_optional=True)
        if cash_not_strip is not None:
            self.cash = self.strip(cash_not_strip)
        else:
            self.cash = self.strip(self.balance.match_title('Cash And Equivalents'))
            assert self.cash is not None

        investment_not_strip = self.balance.match_title('Long-term Investments', none_is_optional=True)
        if investment_not_strip is not None:
            self.investments = self.strip(investment_not_strip)
        else:
            self.investments = 0

        minority_not_strip = self.income.match_title('Minority Interest', none_is_optional=True)
        if minority_not_strip is not None:
            self.minority = self.strip(minority_not_strip)
        else:
            self.minority = 0

        self.shares = self.strip(self.income.match_title('Weighted Average Diluted Shares Outstanding'))
        forward_etr_not_strip = self.trim_estimates('Effective Tax Rate', none_is_optional=True)
        if forward_etr_not_strip is not None:
            self.forward_etr = forward_etr_not_strip
        else:
            self.forward_etr = 0

        self.marginal_tax_rate = .25

        # Malaysia 10 years GBY
        # self.riskfree_rate = .03884

        # U.S. 10 years GBY
        self.riskfree_rate = .0408

    def trim_estimates(self, title, nlead=9, n=4, **args):
        # Remove past annual/quarterly data from Estimates.
        est = self.estimates.match_title(title, **args)
        excess = next((i for i in range(1, n) if est[nlead:][-i] is not None), 0)
        if excess-1 > 0:
            result = est[nlead:][:-excess+1]
        else:
            result = est[nlead:]
        return result

    def compute(self):
        d = OrderedDict()
        self.compute_revenue(d)
        self.compute_ebit(d)
        self.compute_tax(d)
        self.compute_ebt(d)
        self.compute_reinvestment(d)
        self.compute_fcff(d)
        d['empty1'] = []
        self.compute_cost_of_capital(d)
        self.compute_cumulative_df(d)
        self.compute_terminals(d)

        headers = list(d.keys())
        excel = ExcelOut(
            self.tick, d, headers=headers,
            styles=[
                # Revenue growth rate demarcation
                'Percent', 'Comma', 'Percent', 'Comma', 'Percent', 'Comma', 'Comma', 'Comma',
                # Cost of capital demarcation
                '', 'Percent', 'Ratio', 'Comma',
                # Terminal cash flow demarcation
                '', 'Comma', 'Percent', 'Comma', 'Comma', 'Comma', 'Comma', 'Comma', 'Comma', 'Comma', 'Comma', 'Comma', 'Comma', 'Comma', 'Ratio2', 'Ratio2', 'Percent'
            ])
        excel.start()

    def compute_revenue(self, d):
        # Compute past
        # print(self.sales[:-3])

        # Compute forward forecast

        sales_growth_rate = d['Revenue growth rate'] = []
        sales = d['Revenue'] = []
        forward_sales = self.forward_sales[-4:]
        for i in range(1, len(forward_sales)):
            grow_rate = (forward_sales[i] - forward_sales[i-1]) / forward_sales[i-1]
            sales_growth_rate.append(grow_rate)
            sales.append(forward_sales[i])

        # Stable at year 2 and 3
        stable_growth_rate = sales_growth_rate[-1]
        cur_sales = forward_sales[-1] * (1+stable_growth_rate)
        sales_growth_rate.append(stable_growth_rate)
        sales.append(cur_sales)

        cur_sales = cur_sales * (1+stable_growth_rate)
        sales_growth_rate.append(stable_growth_rate)
        sales.append(cur_sales)

        # https://tradingeconomics.com/united-states/government-bond-yield
        # https://tradingeconomics.com/malaysia/government-bond-yield
        # https://tradingeconomics.com/china/government-bond-yield
        # https://tradingeconomics.com/hong-kong/government-bond-yield
        # Terminal year period is based on current risk free rate based on 10 years treasury bond note yield
        term_year_per = self.riskfree_rate

        # Iterating from first year to terminal year in descending grow order, including terminal year
        for n in range(1, 6):
            per = stable_growth_rate - (stable_growth_rate-term_year_per)/5 * n
            sales_growth_rate.append(per)
            cur_sales = cur_sales*(1+per)
            sales.append(cur_sales)

        # Terminal period, no grow and stagnated value
        per = stable_growth_rate - (stable_growth_rate-term_year_per)/5 * 5
        sales_growth_rate.append(per)
        cur_sales = cur_sales * (1 + per)
        sales.append(cur_sales)

    def compute_ebit(self, d):
        sales = d['Revenue']
        ebit_margin = d['EBIT margin'] = []
        ebit = d['EBIT'] = []

        for i, e in enumerate(self.forward_ebit):
            ebit_margin.append(e / sales[i])
            ebit.append(e)
        fixed_index = len(self.forward_ebit)-1

        fixed_margin = ebit_margin[-1]
        for x in range(1, 8):
            ebit_margin.append(fixed_margin)
            stable_ebit = fixed_margin * sales[fixed_index+x]
            ebit.append(stable_ebit)
        ebit_margin.append(fixed_margin)
        ebit.append(fixed_margin * sales[-1])

    def compute_tax(self, d):
        etr = d['Tax rate'] = []
        if self.forward_etr == 0:
            colour_print("Is \"{}\" a REIT company?"
                         " REIT company distributes at least 90% of its total yearly income to unit holders, the REIT itself is exempt from tax for that year, but unit holders are taxed on the distribution of income"
                         .format(self.tick), bcolors.WARNING)
            return

        for i, e in enumerate(self.forward_etr):
            if e is not None:
                etr.append(e/100)
            else:
                # TODO: Invalid ETR
                etr.append(0)

        # Previous year tax rate + (marginal tax rate - previous year tax rate) / 5
        start_tax_rate = tax_rate = etr[-1]
        etr.append(tax_rate)
        etr.append(tax_rate)
        for x in range(1, 6):
            tax_rate = tax_rate+(self.marginal_tax_rate - start_tax_rate)/5
            etr.append(tax_rate)
        etr.append(tax_rate)

    def compute_ebt(self, d):
        ebit = d['EBIT']
        tax_rate = d['Tax rate']
        nopat = d['NOPAT'] = []
        for i, e in enumerate(self.forward_ebit[-3:]):
            nopat.append(e)
            if len(tax_rate) > 0:
                nopat[-1] -= e * tax_rate[i]

        for x in range(0, 8):
            nopat.append(ebit[3+x])
            if len(tax_rate) > 0:
                nopat[-1] -= ebit[3+x] * tax_rate[3+x]

    def compute_reinvestment(self, d):
        # TODO Actual capex, D&A and changes in working capital
        # fwd_capex = self.strip(self.estimates.match_title('Capital Expenditure'))
        # fwd_dna = self.strip(self.estimates.match_title('Depreciation & Amortization'))
        # fwd change in working capital

        reinvestment = d['- Reinvestment'] = []
        sales = d['Revenue']
        for i in range(len(sales)-1):
            # TODO Hard coded 2.5 for sales to capital ratio
            # https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/capex.html
            r = (sales[i+1]-sales[i])/2.5
            reinvestment.append(r)

        # Terminal growth rate / End of ROIC * End of NOPAT
        term_growth_rate = d['Revenue growth rate'][-1]
        # TODO Cost of capital at year 10 or enter manually
        roic = .15
        nopat_end = d['NOPAT'][-1]
        reinvestment.append(term_growth_rate / roic * nopat_end)

    def compute_fcff(self, d):
        fcff = d['FCFF'] = []
        nopat = d['NOPAT']
        reinvestment = d['- Reinvestment']
        for i in range(len(nopat)):
            fcff.append(nopat[i]-reinvestment[i])

    def compute_cost_of_capital(self, d):
        # TODO Cost of capital
        initial_coc = .086

        # Country risk premium set to 4.5% based on U.S. CRP
        # https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html
        country_risk_premium = .045
        coc = d['Cost of capital'] = [initial_coc]*5
        for i in range(1, 6):
            # prev coc - (fixed prev coc - risk free rate + country risk premium)/5
            _ = coc[i-1] - (initial_coc - (self.riskfree_rate + country_risk_premium))/5
            coc.append(_)
        coc.append(self.riskfree_rate + country_risk_premium)

    def compute_cumulative_df(self, d):
        coc = d['Cost of capital']
        fcff = d['FCFF']
        cumulated_df = d['Cumulated discount factor'] = []
        pv = d['PV (FCFF)'] = []
        for i in range(0, 10):
            _ = 1/(1+coc[i])
            if len(cumulated_df) > 0:
                _ = cumulated_df[i-1]*(1/(1+coc[i]))
            cumulated_df.append(_)

            # fcff * df
            pv.append(fcff[i] * cumulated_df[i])

    def compute_terminals(self, d):
        d['empty2'] = []
        d['Terminal cash flow'] = d['FCFF'][-1]
        d['Terminal cost of capital'] = d['Cost of capital'][-1]
        d['Terminal value'] = (d['Terminal cash flow'] /
                               (d['Terminal cost of capital'] - d['Revenue growth rate'][-1]))
        # print( d['Cumulated discount factor'] )
        d['PV (Terminal value)'] = d['Terminal value'] * d['Cumulated discount factor'][-1]
        d['PV (Cash flow over next 10 years)'] = sum(d['PV (FCFF)'])
        d['Sum of PV'] =  d['PV (Terminal value)'] + d['PV (Cash flow over next 10 years)']
        d['Value of operating assets'] =  d['Sum of PV']
        d['- Debt'] = self.debt[-1]
        d['- Minority interest'] = 0
        d['+ Cash'] = self.cash[-1]
        if type(self.investments) is list:
            # type: List[float]
            d['+ Non-operating assets'] = self.investments[-1]
            if self.investments[-1] is None:
                colour_print("Latest investment was not defined. Fallback to previous year", bcolors.WARNING)
                assert self.investments[-2] is not None
                d['+ Non-operating assets'] = self.investments[-2]
        else:
            d['+ Non-operating assets'] = 0
        d['Value of equity'] = (d['Value of operating assets']
                                - d['- Debt'] - d['- Minority interest']
                                + d['+ Cash'] + d['+ Non-operating assets'])
        d['Number of shares'] = self.shares[-1]
        d['Estimated value / share'] = d['Value of equity'] / d['Number of shares']
        d['Price'] = self.strip(self.values.match_title('Price$'))[-1]
        d['Price as % of value'] = d['Price'] / d['Estimated value / share']


class Ticks:
    pass
    # def __init__(self, ticks, path):
    #     self.tickers = ticks
    #     self.path = path
    #     self.modes = [
    #         Mode.Last_FCF,
    #         Mode.Avg_FCF,
    #         Mode.UFCF_5x,
    #         Mode.UFCF_8x,
    #     ]
    #
    #     # type: List[DCF]
    #     self.spreads = []
    #     for t in ticks:
    #         self.spreads.append(DCF(t, self.path, self.modes))
    #
    # def compute(self):
    #     for sp in self.spreads:
    #         sp.compute_fcf()
    #
    # def summarize(self):
    #     # https://blog.devgenius.io/how-to-easily-print-and-format-tables-in-python-18bbe2e59f5f
    #     # summarize to TV based on last FCF, avg FCF, multiple of NOPAT
    #
    #     # Set the first word as the header for our spreadsheet optionally.
    #     a = list(map(lambda x: [x.short_name() if x.head is not None else x.head,
    #                             x.tick, x.last_price], self.spreads))
    #     for i, s in enumerate(self.spreads):
    #         for j in range(len(s.sum_pvtv)):
    #             # extend the spread with sum_pvtv and poss ratio
    #             a[i].append(s.sum_pvtv[j])
    #             a[i].append(s.poss_ratio[j])
    #
    #     # Sort it to the last column which is 'Possible' ratio
    #     entries = sorted(a, key=lambda x: x[len(a[0])-1])
    #     poss_header = 'Poss. x'
    #     heads = ['Company', 'Tick', 'Last price']
    #     for i, m in enumerate(self.modes):
    #         # TODO mapping to the underlying string
    #         h = {Mode.Last_FCF: 'Last FCF',
    #              Mode.Avg_FCF: 'Avg. FCF',
    #              Mode.UFCF_5x: 'UFCF 5x',
    #              Mode.UFCF_8x: 'UFCF 8x'}[m]
    #         heads.extend((h, poss_header))
    #     print(tabulate(entries, headers=heads,
    #                    tablefmt='fancy_grid', stralign='center', numalign='center', floatfmt=".2f"))
    #
    #     # generate Excel output
    #     # styles = ['Comma', 'Comma', 'Comma']
    #     # for _ in self.modes:
    #     #     styles.extend(('Comma', 'Percent'))
    #     # excel = ExcelOut(self.tickers, entries, styles=styles, headers=heads)
    #     # excel.start()


dcf = DCF('intc', 'spreads')
dcf.compute()
print("XXX", dcf)

# Damodaran main data page
# https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datacurrent.html
