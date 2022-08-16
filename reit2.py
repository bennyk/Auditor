from openpyxl import load_workbook, Workbook
from typing import List, Tuple
from collections import OrderedDict, namedtuple
from enum import Enum
import re

max_row = max_col = 99

Unit = namedtuple('Unit', ['value', 'symbol'])
RateType = namedtuple('Rate', ['above_avg', 'moderate_avg', 'below_avg'])

RateVerbose = {RateType.above_avg: 'above',
               RateType.moderate_avg: 'moderately at',

               # TODO "do not perform" below level
               RateType.below_avg: 'below'}


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


def strip2(l, trim_last=False):
    # TODO need to check against original strip implementation.
    if trim_last:
        return l[1:][:-1]

    _ = l[1:]
    if type(_[0]) is str:
        return list(map(lambda x: float(re.sub(r'(\d+)x', r'\1', x)), _))
    return _


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
        last_limit = 0
        try:
            # configure spreadsheet based on number of cols.
            for j in range(2, max_col):
                c0 = "{}{}".format(colnum_string(j), 1)
                self.date_range.append(sheet_ranges[c0].value)
                if re.match(r'LTM$', sheet_ranges[c0].value):
                    Table.col_limit = j + 1
                    break
                last_limit = j+1
        except TypeError:
            Table.col_limit = last_limit

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
            if re.match(reg, _[0].strip()):
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
        self.values = None
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
            elif re.match(r'Values', name):
                self.values = tab
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
        self.profiler.collect(avg_net_debt_over_ebit, Tag.net_debt_over_ebit, ProfMethod.AverageYears)

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

    def ev_over_ebit(self):
        if self.values is None:
            # TODO exception to EV over EBIT
            print("Warning: ev_over_ebit: Missing values tab.")
            return
        ev_over_ebit = strip2(self.values.match_title('LTM Total Enterprise Value / EBIT$'))
        avg_ev_over_ebit = striped_average(ev_over_ebit)
        print("EV over EBIT average {:.2f} ratio for: {}".format(
            avg_ev_over_ebit,
            list(map(lambda x: round(x, 2), ev_over_ebit))
        ))
        self.profiler.collect(avg_ev_over_ebit, Tag.ev_over_ebit, ProfMethod.ReverseRatio)

    # TODO retined earnings pay in advance for one year?
    def retained_earnings_ratio(self):
        _ = strip(self.balance.match_title('Retained Earnings$'), trim_last=True)
        # skip the latest annual for retained earnings only
        retained_earnings = _[:-1]
        # TODO some Company such as IGBREIT does not provide Retained Earnings forecast
        # print("XXX retained earnings", len(retained_earnings), retained_earnings)
        _ = strip(self.income.match_title('Net Income$'), trim_last=True)
        # skip the first year annual for net income only
        net_income = _[1:]
        # print("XXX net income", len(net_income), net_income)
        retention_ratio = list_over_list(retained_earnings, net_income)
        # div_paid = strip(self.cashflow.match_title('Common Dividends Paid'))
        # retention_ratio = list_add_list(net_income, div_paid)

        last_retention_ratio = retention_ratio[-1]
        # retention ratio is measured by carry out from the previous annual report,
        # while adding net income to the current annual report

        avg_retention_ratio = striped_average(retention_ratio)
        print("Retention ratio last {:.2f}, average {:.2f} for: {}".format(
            last_retention_ratio,
            avg_retention_ratio,
            list(map(lambda x: round(x, 2), retention_ratio))
        ))
        self.profiler.collect(last_retention_ratio, Tag.retained_earnings_ratio, ProfMethod.Ratio)

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
        # Negating div payout to positive for the math to work easier
        avg_div_payout_ratio = - striped_average(div_payout_ratio)
        print("Dividend payout ratio at average {:.2f} ratio for: {}".format(
            avg_div_payout_ratio,
            list(map(lambda x: round(x, 2), div_payout_ratio))
        ))
        self.profiler.collect(avg_div_payout_ratio, Tag.dividend_payout_ratio, ProfMethod.Average)

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
    Ratio = 6
    ReverseRatio = 7


ProfVerbose = {ProfMethod.CAGR: 'CAGR',
               ProfMethod.IRR: 'IRR',
               ProfMethod.Average: 'average',
               ProfMethod.AveragePerc: 'percent',
               ProfMethod.AverageYears: 'year',
               ProfMethod.Ratio: 'ratio',
               ProfMethod.ReverseRatio: 'ratio',
               }


class Tag(Enum):
    rev_per_share = 1
    affo_per_share = 2
    nav_per_share = 3
    ROCE = 4
    net_debt_over_ebit = 5
    ebit_margin = 6
    retained_earnings_ratio = 7
    dividend_payout_ratio = 8
    ev_over_ebit = 9


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


class Bucket:
    def __init__(self, name, value, method):
        self.name = name
        self.value = value
        self.method = method
        self.score = None


class ProfManager:
    # TODO report about the underlying rate?
    Rate = {Tag.rev_per_share: {'high': .08, 'mid': .04},
            Tag.affo_per_share: {'high': .3, 'mid': .05},
            Tag.nav_per_share: {'high': .08, 'mid': .05},
            Tag.ROCE: {'high': .08, 'mid': .065},
            Tag.net_debt_over_ebit: {'high': 5., 'mid': 8.},
            Tag.ebit_margin: {'high': .7, 'mid': .6},
            Tag.retained_earnings_ratio: {'high': 5., 'mid': .0},
            Tag.dividend_payout_ratio: {'high': 1.5, 'mid': 1.},
            Tag.ev_over_ebit: {'high': 16., 'mid': 18.},
            }

    def __init__(self):
        self.companies = []     # type: List[Prof]
        self.metric = {}

    def create_folder(self, name):
        prof = Prof(name)
        self.companies.append(prof)
        return prof

    def profile(self):
        for p in self.companies:
            # print(p.name)
            p.profile()
        self.bucketize()
        self.benchmark()

    def benchmark(self):
        # Grading process
        # For each metric based on rev_per_share, affo_per_share, nav_per_share, ...
        #   Rank based on above, moderate and below average.
        #       Collect the company for metric and rank level.
        # Accessible via the hash metric: metric[Tag.rev_per_share][RateType.above_avg]

        met = {}
        for k, v in self.metric.items():
            for v1 in v.values():
                if len(v1) > 0:
                    for x in v1:
                        if x.name not in met:
                            met[x.name] = 0
                        met[x.name] += x.score

        final = {RateType.above_avg: [], RateType.moderate_avg: [], RateType.below_avg: []}
        for k, v in met.items():
            if v >= 4.:
                x = final[RateType.above_avg]
            elif v >= 3.:
                x = final[RateType.moderate_avg]
            else:
                x = final[RateType.below_avg]
            x.append(k)

        print()
        for k, v in final.items():
            if k is RateType.above_avg:
                print("Companies that performed above average are: ", end='')
            elif k is RateType.moderate_avg:
                print("Companies that performed moderate average are: ", end='')
            else:
                print("Companies which are below average are: ", end='')
            print("{} out of {} sample".format(v, len(v)))
        print("Total companies sampled thus far is {}".format(len(self.companies)))

    def bucketize(self):
        # TODO namedtuple?

        for x in Tag:
            self.metric[x] = {RateType.above_avg: [], RateType.moderate_avg: [], RateType.below_avg: [], }

        for c in self.companies:
            for k, v in c.prof.items():
                if type(k) is not str:
                    if k in ProfManager.Rate:
                        assert k in self.metric
                        buck = self.metric[k]
                        tup = Bucket(c.name, v[0], v[1])
                        if v[1] in (ProfMethod.AverageYears, ProfMethod.ReverseRatio):
                            if v[0] < ProfManager.Rate[k]['high']:
                                buck[RateType.above_avg].append(tup)
                            elif v[0] < ProfManager.Rate[k]['mid']:
                                buck[RateType.moderate_avg].append(tup)
                            else:
                                buck[RateType.below_avg].append(tup)
                        else:
                            if v[0] > ProfManager.Rate[k]['high']:
                                buck[RateType.above_avg].append(tup)
                            elif v[0] > ProfManager.Rate[k]['mid']:
                                buck[RateType.moderate_avg].append(tup)
                            else:
                                buck[RateType.below_avg].append(tup)

        def value(val):
            return val.value

        def item(key: Bucket):
            return key.name, key.value

        def articulate(buckets: List[Bucket], key):
            values = list(map(value, buckets))
            if len(values) > 0:
                # TODO need to enumerate the index: -
                #  0 == company,
                #  1 == value of percent, years, etc.,
                #  2 == name of method see ProfMethod class
                method = buckets[0].method
                unit_ratio = (ProfMethod.AverageYears, ProfMethod.Ratio, ProfMethod.ReverseRatio)

                def at(k):
                    if method in unit_ratio:
                        return '{} at {:.2f} yrs'.format(k[0], k[1])
                    return '{} at {:.2f}%'.format(k[0], k[1] * 100)

                items = list(map(item, buckets))
                comp_at_perc = ', '.join(list(map(at, items)))

                unit = Unit(value=100, symbol='%')
                if method in unit_ratio:
                    unit = Unit(value=1, symbol='')

                # TODO modify current "performed above average rate"
                #  to "below the average over the last 10 years sampled, at undemanding rate"
                print("{}/{} companies sampled have performed {} average rate of {} {:.2f}{}. ".format(
                    len(buckets), len(self.companies), RateVerbose[key], ProfVerbose[method], average(values) * unit.value, unit.symbol), end='')
                if key is RateType.above_avg:
                    print("These companies are: {}".format(comp_at_perc))
                    # TODO apply() function?
                    for a in buckets:
                        a.score = 1.
                elif key is RateType.moderate_avg:
                    print("These companies are: {}".format(comp_at_perc))
                    for a in buckets:
                        a.score = .5
                else:
                    for a in buckets:
                        a.score = 0
                    print()

        # TODO need to solve for AFFO, net debt over ebit, retention ratio, div payout ratio

        for k, v in self.metric.items():
            print("Based on {}:".format(k))
            # TODO Near right but not enum hmmm RateType._fields:
            for avg_rate in RateType.above_avg, RateType.moderate_avg, RateType.below_avg:
                articulate(v[avg_rate], avg_rate)


def main():
    path = "C:/Users/benny/iCloudDrive/Documents/malaysia reits"
    prof = ProfManager()

    tickers = ['axreit', 'igbreit', 'sunreit', 'pavreit', 'alaqar',
               'uoareit', 'hektar',
               'klcc', 'amfirst', 'ytlreit', 'arreit', 'clmt',
               'twrreit', 'ahp', 'kipreit',
               'sentral', 'atrium', 'alsreit']

    # tickers = [ 'clmt']
    # tickers = ['igbreit', 'axreit', 'klcc', 'kipreit', 'ytlreit', 'ahp', 'atrium']
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
        t.retained_earnings_ratio()
        t.ebit_margin()
        t.ev_over_ebit()
        t.dividend_payout_ratio()
        print()

    prof.profile()


if __name__ == '__main__':
    main()
