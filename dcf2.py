import pandas as pd

from spread import Spread
from utils import *

from openpyxl import Workbook, load_workbook, worksheet
from openpyxl.styles import Font, Alignment
from collections import OrderedDict

class ExcelOut:
    start_col = 2
    row_margin = 1

    def __init__(self, ticks: [str], entries, styles=None, headers=None):
        cls = self.__class__
        self.od = entries
        self.headers = headers
        self.ticks = ticks
        self.styles = styles
        self.ft = Font(name='Calibri', size=11)
        self.wb = Workbook()

        # type: worksheet.Worksheet
        self.sheet = self.wb.active
        self.sheet.title = 'sheet 1'

        self.cell = self.sheet.cell(row=1, column=cls.start_col)
        self.start_row_index = cls.row_margin+1
        self.end_row_index = len(self.ticks) + self.start_row_index+1

        self.j = cls.row_margin + 1
        self.i = 1

        self.init_sheet()

    def init_sheet(self):
        sheet = self.sheet
        self.j = 1
        i = 2
        for val in self.headers:
            sheet.column_dimensions[colnum_string(self.j)].width = 30
            cell = sheet.cell(row=i, column=self.j)
            cell.alignment = Alignment(wrapText=True)
            cell.value = val
            i += 1

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
            for val in self.od[key]:
                self.make_cell(val, self.styles[i])

            self.j += 1
        self.wb.save('out.xlsx')

    def make_cell(self, e, style):
        cell = self.sheet.cell(row=self.j, column=self.i)
        self.sheet.column_dimensions[colnum_string(self.i)].width = 10
        cell.alignment = Alignment(wrapText=True)
        cell.value = e
        if style == 'Comma':
            cell.style = 'Comma'
            cell.number_format = '0,0'
        else:
            cell.style = 'Percent'
            cell.number_format = '0.00%'
        cell.font = self.ft
        self.i += 1


class DCF(Spread):
    def __init__(self, tick, path):
        self.wb = load_workbook(path + '/' + tick + '.xlsx')
        super().__init__(self.wb, tick)

        # Revenues, Operating Income, Interest Expense, ...

        self.sales = self.strip(self.estimates.match_title('Revenue'))
        self.ebit = self.strip(self.estimates.match_title('EBIT'))
        self.ie = self.strip(self.estimates.match_title('Interest Expense'))
        self.equity = self.strip(self.balance.match_title('Total Equity'))
        self.debt = self.strip(self.balance.match_title('Total Debt'))
        self.cash = self.strip(self.balance.match_title('Total Cash'))
        self.investments = self.strip(self.balance.match_title('Long-term Investments'))
        self.minority = self.strip(self.income.match_title('Minority Interest'))
        self.shares = self.strip(self.income.match_title('Weighted Average Diluted Shares Outstanding'))
        self.eff_tax_rate = self.strip(self.estimates.match_title('Effective Tax Rate'))
        self.marginal_tax_rate = 25.0

    def compute(self):
        d = OrderedDict()
        self.revenue(d)

        headers = list(d.keys())
        excel = ExcelOut(['intc'], d, headers=headers, styles=['Percent', 'Comma'])
        excel.start()

    def revenue(self, d):
        # Compute past
        # print(self.sales[:-3])

        # Compute forward forecast

        sales_growth_rate = d['Revenue growth rate'] = []
        sales = d['Revenue'] = []
        forward_sales = self.sales[-4:]
        for i in range(1, len(forward_sales)):
            grow_rate = (forward_sales[i] - forward_sales[i-1]) / forward_sales[i-1]
            sales_growth_rate.append(grow_rate)
            sales.append(forward_sales[i])

        # Stable at year 2 and 3
        stable_growth_rate = sales_growth_rate[-1]
        cur_sales = forward_sales[-1]  * (1+stable_growth_rate)
        sales_growth_rate.append(stable_growth_rate)
        sales.append(cur_sales)

        cur_sales = cur_sales * (1+stable_growth_rate)
        sales_growth_rate.append(stable_growth_rate)
        sales.append(cur_sales)

        # https://tradingeconomics.com/united-states/government-bond-yield
        # Terminal year period is based on current risk free rate based on 10 years treasury bond note yield
        term_year_per = 0.04

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