from openpyxl import load_workbook, Workbook
from typing import List
import re

max_row = max_col = 99


def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def average(l: [float]):
    l = strip(l)
    return sum(l) / len(l)


def strip(l, trim_last=False):
    if trim_last:
        return l[1:][:-1]
    return l[1:]


def list_over_list(x, y, percent=False):
    if percent:
        # return list(map(lambda n1, n2: (n1 / n2), x, y))
        return list(map(lambda n1, n2: 0 if n1 is None else 100 * (n1 / n2), x, y))
    return list(map(lambda n1, n2: 0 if n1 is None else n1 / n2, x, y))


def list_add_list(x, y):
    a = map(lambda n1: 0 if n1 is None else n1, x)
    b = map(lambda n2: 0 if n2 is None else n2, y)
    return list(map(lambda n1, n2: n1 + n2, a, b))


def list_minus_list(x, y):
    a = map(lambda n1: 0 if n1 is None else n1, x)
    b = map(lambda n2: 0 if n2 is None else n2, y)
    return list(map(lambda n1, n2: n1 - n2, a, b))


def cagr(l: [float]) -> float:
    return (l[len(l) - 1] / l[0]) ** (1 / len(l)) - 1


class Table:
    col_limit = 0

    def __init__(self, sheet_ranges):
        self.date_range = []
        # configure spreadsheet based on number of cols.
        for j in range(2, max_col):
            c0 = "{}{}".format(colnum_string(j), 1)
            self.date_range.append(sheet_ranges[c0].value)
            if re.match(r'LTM$', sheet_ranges[c0].value):
                Table.col_limit = j + 1
                break

        self.tab = []
        for i in range(1, max_row):
            c0 = "{}{}".format(colnum_string(1), i)
            if sheet_ranges[c0].value is None:
                break
            r = []
            for j in range(1, Table.col_limit):
                c1 = "{}{}".format(colnum_string(j), i)
                r.append(sheet_ranges[c1].value)
            self.tab.append(r)

    def match_title(self, reg):
        result = None
        for _ in self.tab:
            if re.match(reg, _[0]):
                result = _
                break
        assert result is not None
        return result


class Spread:
    def __init__(self, wb, tick):
        self.tick = tick
        self.tabs = []
        self.income = None
        self.balance = None
        self.cashflow = None
        for index, name in enumerate(wb.sheetnames):
            tab = Table(wb[name])
            self.tabs.append(tab)
            if re.match(r'Income', name):
                self.income = tab
                self.start_date = self.income.tab[0][1]
                # -2 ignore year LTM
                self.end_date = self.income.tab[0][-2]
            elif re.match(r'Balance', name):
                self.balance = tab
            elif re.match(r'Cash', name):
                self.cashflow = tab
            else:
                # passing Ratios
                pass

        self.end_year = int(self.end_date.split('/')[-1])
        self.start_year = int(self.start_date.split('/')[-1])
        print("Sampled from {} to {} in {} years".format(
            self.start_date, self.end_date,
            1+self.end_year-self.start_year))

    def revenue(self):
        revs = strip(self.income.match_title('Total Revenues'))
        revs_per_share = cagr(list(map(lambda f: f / self.share_out_filing(), revs)))
        print("Revenue per share from {} to {} at CAGR {:.2f}% for: {}".format(
            revs[0], revs[-1],
            revs_per_share * 100, revs))
        if abs(cagr(revs) - revs_per_share) > .01:
            print("   Revenue {:.2f} in percent vs {:.2f} on per share basis".format(
                cagr(revs) * 100, revs_per_share * 100))

    def cfo(self):
        # aka FFO - Funds from Operations
        cfo = strip(self.cashflow.match_title('Cash from Operations'))
        cfo_per_share = cagr(list(map(lambda f: f / self.share_out_filing(), cfo)))
        print("FCF per share from {} to {} at CAGR {:.2f}% for: {}".format(
            cfo[0], cfo[-1],
            cfo_per_share * 100, cfo))
        if abs(cagr(cfo) - cfo_per_share) > .01:
            print("   FCF {:.2f} in percent vs {:.2f} on per share basis".format(
                cagr(cfo) * 100, cfo_per_share * 100))

    def affo(self):
        cfo = strip(self.cashflow.match_title('Cash from Operations'))
        # Capex for real estates
        capex = strip(self.cashflow.match_title('Acquisition of Real Estate Assets'))
        affo = list_add_list(cfo, capex)

        # TODO made comparison in relation to IGBREIT's share out filing
        # share_out_filing = self.share_out_filing()
        share_out_filing = 3600
        affo_per_share = list(map(lambda f: f / share_out_filing, affo))

        # Based on IGBREIT 2021 annual: "term period" between 5.85% to 6.85%. Take the mid point.
        irr = .0635
        term_period = 0

        # reversed - Simulate in reversed order from current to far past year.
        for i, a in enumerate(reversed(affo)):
            term_period += a/(1+irr)**(i+1)
        # print("XXX", term_period, term_period/self.share_out_filing())
        print("AFFO in relation to IGBREIT, at IRR {:.4f} for: {}".format(
            term_period/share_out_filing,
            list(map(lambda x: round(x, 4), affo_per_share))
        ))

    def nav(self):
        total_asset = strip(self.balance.match_title('Total Assets'))
        total_liab = strip(self.balance.match_title('Total Liabilities'))
        nav = list_minus_list(total_asset, total_liab)
        nav_per_share = list(map(lambda f: f / self.share_out_filing(), nav))
        print("NAV per share at CAGR {:.2f}% for: {}".format(
            cagr(nav_per_share)*100,
            list(map(lambda x: round(x, 4), nav_per_share)),
        ))

    def return_equity(self):
        net_income = strip(self.income.match_title('Net Income$'))
        requity = strip(self.balance.match_title('Total Common Equity$'))
        roce = list_over_list(net_income, requity, percent=True)
        print("Return on Common Equity average {:.2f}% for: {}".format(
            average(roce),
            list(map(lambda x: round(x, 2), roce))
        ))

    def net_debt_over_ebit(self):
        net_debt = strip(self.balance.match_title('Net Debt'))
        ebit = strip(self.income.match_title('Operating Income$'))
        # ebitda = self.match_title('EBITDA$')
        print("Net debt over EBIT average {:.2f} years for: {}".format(
            average(list_over_list(net_debt, ebit)),
            list(map(lambda x: round(x, 2), list_over_list(net_debt, ebit)))
        ))

    def ebit_margin(self):
        ebits = strip(self.income.match_title('Operating Income$'))
        revs = strip(self.income.match_title('Total Revenues$'))
        ebit_margins = list_over_list(ebits, revs, percent=True)
        print("EBIT margin average {:.2f}% for (numbers in percent) {}".format(
            average(ebit_margins),
            list(map(lambda x: round(x, 2), ebit_margins))
        ))

    # TODO retined earnings pay in advance for one year?
    def retained_earnings(self):
        retained_earnings = strip(self.balance.match_title('Retained Earnings$'), trim_last=True)
        # TODO some Company such as IGBREIT does not provide Retained Earnings forecast
        # print("XXX retained earnings", len(retained_earnings), retained_earnings)
        net_income = strip(self.income.match_title('Net Income$'), trim_last=True)
        # print("XXX net income", len(net_income), net_income)
        retention_ratio = list_over_list(retained_earnings, net_income)
        # div_paid = strip(self.cashflow.match_title('Common Dividends Paid'))
        # retention_ratio = list_add_list(net_income, div_paid)
        print("Retention ratio last {:.2f}, average {:.2f} for: {}".format(
            retention_ratio[-1],
            average(retention_ratio),
            list(map(lambda x: round(x, 2), retention_ratio))
        ))

    def dividend_payout_ratio(self):
        div_paid = strip(self.cashflow.match_title('Common Dividends Paid'))
        for i, a in enumerate(div_paid):
            if a is None:
                print("W: {} does not provide dividend in year '{}".format(
                    self.tick, self.start_year+i))
        income = strip(self.income.match_title('Net Income'))

        # Op income is a probable replacement in the event when regular income produce negative number.
        op_income = strip(self.income.match_title('Operating Income'))
        net_income = []
        for i, a in enumerate(income):
            if a < 0:
                print("W: {} negative income {} in year '{}".format(
                    self.tick, a, self.start_year+i))
                net_income.append(op_income[i])
            else:
                net_income.append(income[i])

        div_payout_ratio = list_over_list(div_paid, net_income)
        print("Dividend payout ratio at average {:.2f} ratio for: {}".format(
            average(div_payout_ratio),
            list(map(lambda x: round(x, 2), div_payout_ratio))
        ))

    def share_out_filing(self) -> float:
        x = self.balance.match_title('Total Shares Out\.')
        result = list(filter(None, reversed(x[1:])))[0]
        return result


def main():
    path = "C:/Users/benny/iCloudDrive/Documents/malaysia reits"

    tickers = ['axreit', 'igbreit', 'sunreit', 'pavreit', 'alaqar',
               'uoareit', 'hektar',
               'klcc', 'amfirst', 'ytlreit', 'arreit', 'clmt',
               'twrreit', 'ahp', 'kipreit']

    # tickers = [ 'atrium']
    # tickers = [ 'clmt']
    # tickers = [ 'axreit']
    # tickers = ['igbreit', 'axreit', 'klcc', 'kipreit', 'ahp']
    for c in tickers:
        print('Ticker {}'.format(c))
        wb = load_workbook(path+'/' + c + '.xlsx')
        t = Spread(wb, c)
        t.revenue()
        # t.cfo()
        t.affo()
        t.nav()
        t.return_equity()
        t.net_debt_over_ebit()
        t.retained_earnings()
        t.ebit_margin()
        t.dividend_payout_ratio()
        print()


if __name__ == '__main__':
    main()
