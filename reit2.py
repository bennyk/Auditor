from openpyxl import load_workbook, Workbook
from typing import List, Tuple
from collections import OrderedDict, namedtuple
from enum import Enum
import re

max_row = max_col = 99


def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def striped_average(l: [float]):
    l = strip(l)
    return sum(l) / len(l)


def average(l: [float]):
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
    def __init__(self, wb, tick, prof: 'Prof'):
        self.tick = tick
        self.profiler = prof
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
        rev_per_share = cagr(list(map(lambda f: f / self.share_out_filing(), revs)))
        print("Revenue per share from {} to {} at CAGR {:.2f}% for: {}".format(
            revs[0], revs[-1],
            rev_per_share * 100, revs))
        if abs(cagr(revs) - rev_per_share) > .01:
            print("   Revenue {:.2f} in percent vs {:.2f} on per share basis".format(
                rev_per_share * 100, rev_per_share * 100))

        self.profiler.collect(rev_per_share, Tag.rev_per_share, ProfMethod.CAGR)

    def cfo(self):
        # aka FFO - Funds from Operations
        cfo = strip(self.cashflow.match_title('Cash from Operations'))
        cfo_per_share = cagr(list(map(lambda f: f / self.share_out_filing(), cfo)))
        print("FCF per share from {} to {} at CAGR {:.2f}% for: {}".format(
            cfo[0], cfo[-1],
            cfo_per_share * 100, cfo))
        if abs(cagr(cfo) - cfo_per_share) > .01:
            print("   FCF {:.2f} in percent vs {:.2f} on per share basis".format(
                cfo_per_share*100, cfo_per_share * 100))
        self.profiler.collect(cfo_per_share, 'cfo_per_share', ProfMethod.CAGR)

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
        avg_term_period_over_shares = term_period / share_out_filing
        print("AFFO in relation to IGBREIT, at IRR {:.4f} for: {}".format(
            avg_term_period_over_shares,
            list(map(lambda x: round(x, 4), affo_per_share))
        ))
        self.profiler.collect(avg_term_period_over_shares, Tag.affo_per_share, ProfMethod.IRR)

    def nav(self):
        total_asset = strip(self.balance.match_title('Total Assets'))
        total_liab = strip(self.balance.match_title('Total Liabilities'))
        nav = list_minus_list(total_asset, total_liab)
        nav_per_share = list(map(lambda f: f / self.share_out_filing(), nav))
        avg_nav_per_share = cagr(nav_per_share)
        print("NAV per share at CAGR {:.2f}% for: {}".format(
            avg_nav_per_share*100,
            list(map(lambda x: round(x, 4), nav_per_share)),
        ))
        self.profiler.collect(avg_nav_per_share, Tag.nav_per_share, ProfMethod.CAGR)

    def return_equity(self):
        net_income = strip(self.income.match_title('Net Income$'))
        requity = strip(self.balance.match_title('Total Common Equity$'))
        roce = list_over_list(net_income, requity, percent=True)
        avg_roce = striped_average(roce)
        print("Return on Common Equity average {:.2f}% for: {}".format(
            avg_roce,
            list(map(lambda x: round(x, 2), roce))
        ))
        # TODO ROCE in percent
        self.profiler.collect(avg_roce/100, Tag.ROCE, ProfMethod.AveragePerc)

    def net_debt_over_ebit(self):
        net_debt = strip(self.balance.match_title('Net Debt'))
        ebit = strip(self.income.match_title('Operating Income$'))
        # ebitda = self.match_title('EBITDA$')
        net_debt_over_ebit = list_over_list(net_debt, ebit)
        avg_net_debt_over_ebit = striped_average(net_debt_over_ebit)
        print("Net debt over EBIT average {:.2f} years for: {}".format(
            avg_net_debt_over_ebit,
            list(map(lambda x: round(x, 2), net_debt_over_ebit))
        ))
        self.profiler.collect(avg_net_debt_over_ebit, 'net_debt_over_ebit', ProfMethod.AverageYears)

    def ebit_margin(self):
        ebits = strip(self.income.match_title('Operating Income$'))
        revs = strip(self.income.match_title('Total Revenues$'))
        ebit_margins = list_over_list(ebits, revs, percent=True)
        avg_ebit_margins = striped_average(ebit_margins)
        print("EBIT margin average {:.2f}% for (numbers in percent) {}".format(
            avg_ebit_margins,
            list(map(lambda x: round(x, 2), ebit_margins))
        ))
        self.profiler.collect(avg_ebit_margins/100, Tag.ebit_margin, ProfMethod.AveragePerc)

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
        avg_retention_ratio = striped_average(retention_ratio)
        print("Retention ratio last {:.2f}, average {:.2f} for: {}".format(
            retention_ratio[-1],
            avg_retention_ratio,
            list(map(lambda x: round(x, 2), retention_ratio))
        ))
        self.profiler.collect(avg_retention_ratio, 'retention_ratio', ProfMethod.Average)

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
        avg_div_payout_ratio = striped_average(div_payout_ratio)
        print("Dividend payout ratio at average {:.2f} ratio for: {}".format(
            avg_div_payout_ratio,
            list(map(lambda x: round(x, 2), div_payout_ratio))
        ))
        self.profiler.collect(avg_div_payout_ratio, 'dividend_payout_ratio', ProfMethod.Average)

    def share_out_filing(self) -> float:
        x = self.balance.match_title('Total Shares Out\.')
        result = list(filter(None, reversed(x[1:])))[0]
        return result


class ProfMethod(Enum):
    CAGR = 1
    IRR = 2
    Average = 3
    AveragePerc = 4
    AverageYears = 5


class Tag(Enum):
    rev_per_share = 1
    affo_per_share = 2
    nav_per_share = 3
    ROCE = 4
    # net_debt_over_ebit = 5
    ebit_margin = 6


class Prof:
    def __init__(self, name):
        self.name = name
        self.d = OrderedDict()

        self.sps = None
        self.prof = {}

    def collect(self, val, tag, method: ProfMethod):
        # self.d[tag.__name__] = ratio
        # type: Tuple[float, ProfMethod]
        _ = (val, method)
        self.d[tag] = _

    def profile(self):
        for k, v in self.d.items():
            if k is not str:
                self.prof[k] = v


class ProfManager:
    # TODO report about the underlying rate?
    Rate = {Tag.rev_per_share: {'high': .08, 'mid': .04},
            Tag.affo_per_share: {'high': .3, 'mid': .05},
            Tag.nav_per_share: {'high': .08, 'mid': .05},
            Tag.ROCE: {'high': .08, 'mid': .065},
            # Tag.net_debt_over_ebit: {'high': .08, 'mid': .065},
            Tag.ebit_margin: {'high': .7, 'mid': .6},
            }

    def __init__(self):
        self.companies = []     # type: List[Prof]

    def create_folder(self, name):
        prof = Prof(name)
        self.companies.append(prof)
        return prof

    def profile(self):
        for p in self.companies:
            # print(p.name)
            p.profile()
        self.bucketize()

    def bucketize(self):
        # TODO namedtuple?

        metric = {}
        for x in Tag:
            metric[x] = {'above_avg': [], 'moderate_avg': [], 'below_avg': [], }

        for c in self.companies:
            for k, v in c.prof.items():
                if type(k) is not str:
                    if k in ProfManager.Rate:
                        assert k in metric
                        buck = metric[k]
                        tup = (c.name, v[0], v[1])
                        # TODO net_debt_over_ebit need bucketize debt
                        if v[0] > ProfManager.Rate[k]['high']:
                            buck['above_avg'].append(tup)
                        elif v[0] > ProfManager.Rate[k]['mid']:
                            buck['moderate_avg'].append(tup)
                        else:
                            buck['below_avg'].append(tup)

        def value(val):
            # Ignore profile method in v[1][1]
            return val[1]

        def item(key):
            return key[0], key[1]

        def at(key):
            return '{} at {:.2f}%'.format(key[0], key[1] * 100)

        def articulate(bucket, key):
            values = list(map(value, bucket))
            items = list(map(item, bucket))
            comp_at_perc = ', '.join(list(map(at, items)))

            if len(values) > 0:

                # TODO 0 == ticker, 1 == ratio, 2 == CAGR/IRR/years method
                method = bucket[0][2]
                if method is ProfMethod.CAGR:
                    method = 'CAGR'
                elif method is ProfMethod.AveragePerc:
                    method = 'average percent'
                elif method is ProfMethod.IRR:
                    method = 'IRR'

                print("{}/{} companies sampled have performed above average rate at {} {:.2f}%. ".format(
                    len(bucket), len(self.companies), method, average(values)*100), end='')
                if key == 'above_avg':
                    print("These companies are: {}".format(comp_at_perc))
                elif key == 'moderate_avg':
                    print("These companies are: {}".format(comp_at_perc))
                else:
                    print()

        # TODO need to solve for AFFO, net debt over ebit, retention ratio, div payout ratio

        for k, v in metric.items():
            print("Based on {}:".format(k))
            for avg_rate in 'above_avg', 'moderate_avg', 'below_avg':
                articulate(v[avg_rate], avg_rate)


def main():
    path = "C:/Users/benny/iCloudDrive/Documents/malaysia reits"
    prof = ProfManager()

    tickers = ['axreit', 'igbreit', 'sunreit', 'pavreit', 'alaqar',
               'uoareit', 'hektar',
               'klcc', 'amfirst', 'ytlreit', 'arreit', 'clmt',
               'twrreit', 'ahp', 'kipreit']

    # tickers = [ 'atrium']
    # tickers = [ 'clmt']
    # tickers = ['igbreit', 'axreit', 'klcc', 'kipreit', 'ytlreit', 'ahp']
    for c in tickers:
        print('Ticker {}'.format(c))
        wb = load_workbook(path+'/' + c + '.xlsx')
        pf = prof.create_folder(c)
        t = Spread(wb, c, pf)
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

    prof.profile()


if __name__ == '__main__':
    main()
