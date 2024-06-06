import re
import pandas as pd

from spread import Spread
from utils import *
from bcolors import colour_print, bcolors

from openpyxl import Workbook, load_workbook, worksheet
from openpyxl.styles import Font, Alignment
from collections import OrderedDict
from typing import List
import yfinance as yf
from tabulate import tabulate
from calculator import ExcelWriter

total_main_col = 12
total_half_col = int(total_main_col/2)


class DataSet:
    def __init__(self, country, industry):
        self.path = "datacurrent"
        self.country = country
        self.industry = industry

    def get_country_tax_rates(self):
        name = 'countrytaxrates'
        result = None
        wb = load_workbook(self.path + '/{}.xlsx'.format(name))
        ws = wb[name]
        for i in range(1, ws.max_row):
            if ws['A{}'.format(i)].value == self.country:
                result = ws.cell(row=i, column=ws.max_column).value
                break
        assert result is not None
        return result

    def get_riskfree_rate(self):
        # TODO Country 10 years GBY on 13-Apr-2024 at the following websites: -
        # https://tradingeconomics.com/united-states/government-bond-yield
        # https://tradingeconomics.com/malaysia/government-bond-yield
        # https://tradingeconomics.com/china/government-bond-yield
        # https://tradingeconomics.com/hong-kong/government-bond-yield
        # https://tradingeconomics.com/taiwan/government-bond-yield
        tab = {
            'Malaysia': .03947,
            # 'United States': .0408,
            'United States': .04532,
            'Taiwan': .01565,
            'China': .02291,
            'Hong Kong': .0388, }
        assert self.country in tab
        return tab[self.country]

    def get_wacc(self):
        name = 'wacc'
        result = None
        wb = load_workbook(self.path + '/{}.xlsx'.format(name))
        ws = wb['Industry Averages']
        for i in range(20, ws.max_row):
            print(ws['A{}'.format(i)].value)
            if re.match(self.industry, ws['A{}'.format(i)].value, re.IGNORECASE):
                result = ws.cell(row=i, column=ws.max_column).value
                break
        assert result is not None
        return result

    def get_equity_risk_premium(self):
        result = None
        wb = load_workbook(self.path + '/{}.xlsx'.format('ERPs by country'))
        ws = wb['Sheet1']
        for i in range(8, ws.max_row):
            if re.match(self.country, ws['A{}'.format(i)].value):
                result = ws.cell(row=i, column=5).value
                break
        assert result is not None
        return result

    def get_currency_suffix(self, sticky_price):
        assert self.country is not None
        result = None
        with open(self.path + '/yahoo_currency_suffix.txt', encoding='utf-8') as infile:
            # https://www.gnucash.org/docs/v4/C/gnucash-help/fq-spec-yahoo.html
            suffix_index = None
            for i, field in enumerate(infile.readline().split('|')):
                if re.match('suffix', field, re.IGNORECASE):
                    suffix_index = i
                    break
            assert suffix_index is not None
            for line in infile:
                a = line[:-1].split('|')
                if re.match(a[suffix_index][1:], sticky_price):
                    result = a[suffix_index][1:]
                    break
        return result

    def get_sales_to_cap_ratio(self):
        name = 'capex'
        sales_to_cap_index = None
        result = None
        wb = load_workbook(self.path + '/{}.xlsx'.format(name))
        ws = wb['Industry Averages']
        for i in range(1, ws.max_column+1):
            a = ws.cell(row=8, column=i).value
            if re.match(r'Sales/ Invested Capital', a):
                sales_to_cap_index = i
                break
        assert sales_to_cap_index is not None
        for i in range(8, ws.max_row):
            if re.match(self.industry, ws['A{}'.format(i)].value, re.IGNORECASE):
                result = ws.cell(row=i, column=sales_to_cap_index).value
                break
        assert result is not None
        return result

    def match_sales_to_cap_ratio(self, sales_to_cap_ratio):
        name = 'capex'
        sales_to_cap_index = None
        result = None
        wb = load_workbook(self.path + '/{}.xlsx'.format(name))
        ws = wb['Industry Averages']
        for i in range(1, ws.max_column + 1):
            a = ws.cell(row=8, column=i).value
            if re.match(r'Sales/ Invested Capital', a):
                sales_to_cap_index = i
                break
        assert sales_to_cap_index is not None
        result = []
        for i in range(9, ws.max_row):
            diff = ws.cell(row=i, column=sales_to_cap_index).value - sales_to_cap_ratio
            # print(ws.cell(row=i, column=1).value, diff)
            result.append([
                ws.cell(row=i, column=1).value,
                ws.cell(row=i, column=sales_to_cap_index).value, abs(diff), ])
        return sorted(result, key=lambda e: e[2])


class DCF(Spread):
    def __init__(self, tick, country=None, industry=None, path=None):
        colour_print("Company's ticker '{}'".format(tick), bcolors.UNDERLINE)

        if country is None:
            country = 'United States'
            colour_print("Defaulting country to '{}'?".format(country), bcolors.WARNING)

        if industry is None:
            industry = 'semiconductor'
            colour_print("Defaulting industry to '{}'?".format(industry), bcolors.WARNING)

        if path is None:
            path = 'spreads'
            print("Set path to '{}'".format(path))

        self.wb = load_workbook(path + '/' + tick + '.xlsx')
        super().__init__(self.wb, tick)

        self.excel = ExcelWriter(tick)
        self.dataset = DataSet(country, industry)
        self.cached_ticker = None

        # Revenues, Operating Income, Interest Expense, ...

        self.sales = self.strip(self.income.match_title('Total Revenues'))
        self.forward_sales = self.trim_estimates('Revenue', nlead=8)
        self.forward_ebit = self.trim_estimates('EBIT$')

        # TODO Interest expense and Equity?
        self.ie = self.strip(self.income.match_title('Interest Expense'))
        self.book_value_equity = self.strip(self.balance.match_title('Total Equity'))
        self.book_value_debt = self.strip(self.balance.match_title('Total Debt'))
        self.current_debt = [0.]
        current_debt_not_strip = self.balance.match_title('Current Portion of Long-Term Debt', none_is_optional=True)
        if current_debt_not_strip is not None:
           self.current_debt = self.strip(current_debt_not_strip)
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

        self.marginal_tax_rate = self.dataset.get_country_tax_rates()

        # Country 10 years GBY
        self.riskfree_rate = self.dataset.get_riskfree_rate()

        # Sea of change by Howard Marks
        # https://www.oaktreecapital.com/insights/memo/sea-change
        # self.riskfree_rate = .036

    def trim_estimates(self, title, nlead=8, n=4, **args):
        # Remove past annual/quarterly data from Estimates.
        est = self.estimates.match_title(title, **args)
        result = None
        if est is not None:
            excess = next((i for i in range(1, n) if est[nlead:][-i] is not None), 0)
            if excess-1 > 0:
                result = est[nlead:][:-excess+1]
            else:
                result = est[nlead:]
        return result

    def compute(self):
        # d = OrderedDict()
        d = self.excel.create_dict()
        self.compute_revenue(d)
        self.compute_ebit(d)
        self.compute_tax(d)
        self.compute_ebt(d)
        self.compute_reinvestment(d)
        self.compute_fcff(d)
        # d['empty1'] = []
        self.compute_cost_of_capital(d)
        self.compute_cumulative_df(d)
        self.compute_terminals(d)
        self.compute_return_invested_capital(d)

        self.excel.wb.save('aaa.xlsx')

    def compute_revenue(self, d):
        # Compute past
        # Compute forward forecast

        sales_grate_row = 2
        sales_growth_rate = d.create_array('Revenue growth rate', sales_grate_row, style='Percent')
        sales = d.create_array('Revenue', 3)
        forward_sales = self.forward_sales[-4:]
        for i in range(len(forward_sales)):
            # grow_rate = (forward_sales[i] - forward_sales[i-1]) / forward_sales[i-1]
            if i != 0:
                grow_rate_cell = "=({}{row}-{}{row})/{}{row}".format(
                    colnum_string(i+2), colnum_string(i+1), colnum_string(i+1), row=3)
                sales_growth_rate.append(grow_rate_cell)
            else:
                sales_growth_rate.append(0)
            sales.append(forward_sales[i])

        # Stable at year 2, 3 (optional here after), 4 and 5
        # stable_growth_rate = sales_growth_rate[-1]
        # current_sales = forward_sales[-1] * (1+stable_growth_rate)
        stable_growth_rate = "={}{row}".format(
            colnum_string(len(forward_sales)+1), row=sales_grate_row)
        current_sales = "={}{sales_row}*(1+{}{sales_grate_row})".format(
            colnum_string(len(forward_sales)+1), colnum_string(len(forward_sales)+2),
            sales_row=3, sales_grate_row=sales_grate_row)
        for i in range(len(forward_sales)-1, 5):
            sales_growth_rate.append(stable_growth_rate)
            stable_growth_rate = "={}{row}".format(
                colnum_string(len(forward_sales)+1), row=sales_grate_row)
            sales.append(current_sales)
            if i < 4:
                # Otherwise, ignore the last loop
                current_sales = "={}{sales_row}*(1+{}{sales_grate_row})".format(
                    colnum_string(len(forward_sales)+i-1), colnum_string(len(forward_sales)+i),
                    sales_row=3, sales_grate_row=sales_grate_row)
                # current_sales = current_sales * (1+stable_growth_rate)

        # Terminal year period is based on current risk free rate based on 10 years treasury bond note yield
        term_year_per = self.riskfree_rate

        # Iterating from first year to terminal year in descending grow order, including terminal year
        for n in range(1, total_half_col):
            # per = stable_growth_rate - (stable_growth_rate-term_year_per)/5 * n
            per = "={fwd_sales_plus1}{row} - ({fwd_sales_plus1}{row}-{term_year_per})/5 * {n}".format(
                fwd_sales_plus1=colnum_string(6+n), n=n,
                row=sales_grate_row, term_year_per=term_year_per)
            sales_growth_rate.append(per)
            # current_sales = current_sales*(1+per)
            current_sales = "={}{sales_row}*(1+{}{sales_grate_row})".format(
                colnum_string(6+n), colnum_string(7+n),
                sales_row=3, sales_grate_row=sales_grate_row)
            sales.append(current_sales)

        # Terminal period, no grow and stagnated value
        per = "={fwd_sales_plus1}{row} - ({fwd_sales_plus1}{row}-{term_year_per})/5 * 5".format(
            fwd_sales_plus1=colnum_string(12),
            row=sales_grate_row, term_year_per=term_year_per)
        # per = stable_growth_rate - (stable_growth_rate-term_year_per)/5 * 5
        sales_growth_rate.append(per)
        # current_sales = current_sales * (1 + per)
        current_sales = "={}{sales_row}*(1+{}{sales_grate_row})".format(
            colnum_string(12), colnum_string(13),
            sales_row=3, sales_grate_row=sales_grate_row)
        sales.append(current_sales)
        pass

    def compute_ebit(self, d):
        # sales = d['Revenue']
        sales_row = 3
        ebit_margin_row = 4
        ebit_row = 5
        ebit_margin = d.create_array('EBIT margin', ebit_margin_row, style='Percent')
        ebit = d.create_array('EBIT', ebit_row)

        for i, e in enumerate(self.forward_ebit):
            margin_template = "={}{ebit_row} / {}{sales_row}".format(
                colnum_string(i+2), colnum_string(i+2),
                ebit_row=ebit_row, sales_row=sales_row
            )
            # ebit_margin.append(e / sales[i])
            ebit_margin.append(margin_template)
            ebit.append(e)

        fixed_index = len(self.forward_ebit)-1
        # fixed_margin = ebit_margin[-1]
        fixed_margin = '={}{ebit_margin_row}'.format(
            colnum_string(5), ebit_margin_row=ebit_margin_row)
        remaining_col = total_main_col - len(self.forward_ebit)
        for i in range(1, remaining_col):
            ebit_margin.append(fixed_margin)
            # stable_ebit = fixed_margin * sales[fixed_index+i]
            stable_ebit = "={col}{ebit_margin_row}*{col}{sales_row}".format(
                col=colnum_string(i+5), ebit_margin_row=ebit_margin_row, sales_row=sales_row
            )
            ebit.append(stable_ebit)
        ebit_margin.append(fixed_margin)
        # ebit.append(fixed_margin * sales[-1])
        ebit_formula = "={col}{ebit_margin_row}*{col}{sales_row}".format(
            ebit_margin_row=ebit_margin_row, sales_row=sales_row,
            col=colnum_string(13))
        ebit.append(ebit_formula)

    def compute_tax(self, d):
        tax_row = 6
        tax_col = colnum_string(7)
        # etr = d['Tax rate'] = []
        etr = d.create_array('Tax rate', tax_row, style='Percent')
        if self.forward_etr == 0:
            colour_print("Is \"{}\" a REIT company?"
                         " REIT company distributes at least 90% of its total yearly income to unit holders, the REIT itself is exempt from tax for that year, but unit holders are taxed on the distribution of income"
                         .format(self.tick), bcolors.WARNING)
            return

        assert type(self.forward_etr) is list
        for i, e in enumerate(self.forward_etr):
            if e is not None:
                etr.append(e/100)
            else:
                # TODO: Invalid ETR
                etr.append(0)

        # Previous year tax rate + (marginal tax rate - previous year tax rate) / 5
        # start_tax_rate = tax_rate = etr[-1]
        start_tax_rate = "{}{}".format(colnum_string(i+2), tax_row)
        for i in range(len(self.forward_etr), total_half_col):
            # etr.append(tax_rate)
            etr.append("={}".format(start_tax_rate))

        tax_cell = "${}${}".format(tax_col, tax_row)
        for i in range(1, total_half_col):
            tax_rate_cell = "={}{}+({}-{})/5".format(
                colnum_string(i+6), tax_row, self.marginal_tax_rate, tax_cell)
            # tax_rate = tax_rate+(self.marginal_tax_rate - start_tax_rate)/5
            # etr.append(tax_rate)
            etr.append(tax_rate_cell)
        # etr.append(tax_rate)
        tax_rate_cell = "={}{}".format(colnum_string(i+7), tax_row)
        etr.append(tax_rate_cell)

    def compute_ebt(self, d):
        nopat_row = 7
        nopat = d.create_array('NOPAT', nopat_row)
        ebit_row = 5
        tax_row = 6
        # ebit = d['EBIT']
        # tax_rate = d['Tax rate']
        # nopat = d['NOPAT'] = []
        # for i, e in enumerate(self.forward_ebit):
        #     nopat.append(e)
        #     if len(tax_rate) > 0:
        #         nopat[-1] -= e * tax_rate[i]
        for i in range(len(self.forward_ebit)):
            nopat_cell = "={}{}*(1-{}{})".format(colnum_string(i+2), ebit_row,  colnum_string(i+2), tax_row)
            nopat.append(nopat_cell)

        # nopat_start = len(self.forward_ebit)
        # for i in range(nopat_start, total_main_col):
        #     nopat.append(ebit[i])
        #     if len(tax_rate) > 0:
        #         nopat[-1] -= ebit[i] * tax_rate[i]

        nopat_start = len(self.forward_ebit)
        for i in range(nopat_start, total_main_col):
            nopat_cell = "={}{}*(1-{}{})".format(colnum_string(i + 2), ebit_row, colnum_string(i + 2), tax_row)
            nopat.append(nopat_cell)

    def compute_reinvestment(self, d):
        # TODO Actual capex, D&A and changes in working capital
        # fwd_capex = self.strip(self.estimates.match_title('Capital Expenditure'))
        # fwd_dna = self.strip(self.estimates.match_title('Depreciation & Amortization'))
        # fwd change in working capital

        # Defaulting to no lag for the time being.
        # =IF(more_lag="No",(forward_2y_sales-forward_sales)/sales_to_cap,
        #   IF(years_lag=0,(forward_sales-prev_sales)/sales_to_cap,
        #   IF(years_lag=2,(forward_3y_sales-forward_2y_sales)/sales_to_cap,
        #   IF(years_lag=3,(forward_4y_sales-forward_3y_sales)/sales_to_cap,
        #   (forward_2y_sales-forward_sales)/sales_to_cap))))
        # https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/capex.html

        # Latest sales to book value of equity
        sales_to_cap_source = self.sales[-1] / self.book_value_equity[-1]
        print("Computed Sales to cap ratio {:.2f}".format(sales_to_cap_source))
        # TODO Asia countries not in U.S. coverage
        if True:
            print("Probable Sales to cap ratio:")
            matches = self.dataset.match_sales_to_cap_ratio(sales_to_cap_source)[:5]
            heads = ['Company', 'Sales to Cap', 'Error']
            print(tabulate(matches, headers=heads, floatfmt=".2f"), "\n")
            # Selecting mid of the 5 matches
            sales_to_cap_ratio = matches[2][1]
        else:
            sales_to_cap_ratio = sales_to_cap_source

        sales_row = 3
        reinvest_row = 8
        reinvestment = d.create_array('- Reinvestment', reinvest_row)
        # reinvestment = d['- Reinvestment'] = []
        reinvestment.append(None)
        # sales = d['Revenue']
        # for i in range(1, len(sales)-1):
        for i in range(1, 11):
            # r = (sales[i+1]-sales[i]) / sales_to_cap_ratio
            r = "=({}{}-{}{})/{}".format(
                colnum_string(i+3), sales_row, colnum_string(i+2), sales_row, sales_to_cap_source)
            reinvestment.append(r)
        # self.excel.wb.save('aaa.xlsx')
        # exit()

        # Terminal growth rate / End of ROIC * End of NOPAT
        term_growth_rate = "{}".format(d.get('Revenue growth rate').last())
        # term_growth_rate = d['Revenue growth rate'][-1]
        # TODO Cost of capital at year 10 or enter manually
        roic = .15
        # nopat_end = d['NOPAT'][-1]
        # nopat_end = d.get('NOPAT').last()
        terminal_col = 13
        nopat_row = 7
        nopat_cell = '{}{}'.format(colnum_string(terminal_col), nopat_row)
        # current_re = d.concat('{} / {} * {}'.format(term_growth_rate, roic, nopat_cell))
        # reinvestment.append(term_growth_rate / roic * nopat_end)
        reinvestment.append('={}/{} * {}'.format(term_growth_rate, roic, nopat_cell))

        # Alternatively Sales to IC ratio = FCFF (computed by TIKR) + NOPAT

    def compute_fcff(self, d):
        fcff_row = 9
        fcff = d.create_array('FCFF', fcff_row)
        nopat_row = 7
        reinvestment_row = 8
        fcff.append(None)
        for i in range(2, 13):
            fcff_cell = "={}{}-{}{}".format(
                colnum_string(i+1), nopat_row, colnum_string(i+1), reinvestment_row)
            fcff.append(fcff_cell)

        # fcff = d['FCFF'] = []
        # nopat = d['NOPAT']
        # reinvestment = d['- Reinvestment']
        # for i in range(len(nopat)):
        #     if reinvestment[i] is not None:
        #         fcff.append(nopat[i]-reinvestment[i])
        #     else:
        #         fcff.append(None)

    def compute_cost_of_capital(self, d):
        # Cost of debt
        interest_expense = 0
        if self.ie[-1] is not None:
            interest_expense = self.ie[-1]

        debt = 0
        if self.debt[-1] is not None:
            debt = self.debt[-1]
        # I have excluded tax rate leading to lower debt number.
        pretax_cost_of_debt = 0
        if debt > 0:
            pretax_cost_of_debt = abs(interest_expense / debt)
        if self.forward_etr == 0:
            cost_of_debt_formula = pretax_cost_of_debt
        else:
            cost_of_debt_formula = "{}*(1-{}/100)".format(
                pretax_cost_of_debt, average(self.forward_etr))

        # Cost of equity
        yf_ticker = self.get_ticker()
        if 'beta' in yf_ticker.info:
            beta = yf_ticker.info['beta']
            print("Obtain beta:", beta)
        else:
            colour_print("Invalid beta: defaulting beta to 1.0", bcolors.WARNING)
            beta = 1.0
        mrp = self.dataset.get_equity_risk_premium()
        # cost_of_equity = self.riskfree_rate + beta * mrp
        cost_of_equity_formula = "({}+{}*{})".format(self.riskfree_rate, beta, mrp)

        market_cap = yf_ticker.info['marketCap'] / 1e6
        # total_cap = market_cap + debt
        # initial_coc = market_cap/total_cap * cost_of_equity + debt/total_cap * cost_of_debt
        total_cap_formula = "{}+{}".format(market_cap, debt)
        initial_coc_formula = "{}/({})*{} + {}/({})*{}".format(
            market_cap, total_cap_formula, cost_of_equity_formula,
            debt, total_cap_formula, cost_of_debt_formula)

        # Mature market ERP set to 4.5% and 0% based on U.S. CRP (Country risk premium)
        # https://www.youtube.com/watch?v=kyKfJ_7-mdg
        # https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datafile/ctryprem.html
        crp = 0
        mature_market_erp = .045
        coc_row = 11
        coc = d.create_array('Cost of capital', coc_row, style="Percent")
        coc.append(None)
        coc.append("={}".format(initial_coc_formula))
        for j in range(1, 5):
            coc.append("={}{}".format(colnum_string(j+2), coc_row))
        # coc = d['Cost of capital'] = [0] + [initial_coc]*5
        last_coc = "${}${}".format(colnum_string(j+3), coc_row)

        total_erp = crp + mature_market_erp
        for i in range(1, total_half_col):
            # Prev coc - (fixed prev coc - riskfree rate + mature market risk + country risk premium)/5
            # current_coc = coc[i] - (initial_coc - (self.riskfree_rate + total_erp))/5
            current_coc = "={} - ({}-({}+{}))/5".format(last_coc, initial_coc_formula, self.riskfree_rate, total_erp)
            coc.append(current_coc)
        coc.append("={}+{}".format(self.riskfree_rate, mature_market_erp))

    def compute_cumulative_df(self, d):
        coc_row = 11
        fcff_row = 9
        # coc = d['Cost of capital']
        # fcff = d['FCFF']
        # cumulated_df = d['Cumulated discount factor'] = []
        # pv = d['PV (FCFF)'] = []
        cumulated_df_row = 12
        pv_row = 13
        cumulated_df = d.create_array('Cumulated discount factor', cumulated_df_row)
        cumulated_df.append(1.0)
        pv = d.create_array('PV (FCFF)', pv_row)
        pv.append(None)
        for i in range(total_main_col-2):
            current_coc = 0
            # if coc[i] is not None:
            #     current_coc = coc[i]
            # current_df = 1/(1+current_coc)
            # if len(cumulated_df) > 0:
            #     current_df = cumulated_df[i-1]*(1/(1+current_coc))
            current_df = "={}{}*1/(1+{}{})".format(
                colnum_string(i+2), cumulated_df_row, colnum_string(i+3), coc_row)
            cumulated_df.append(current_df)

            current_fcff = "={}{}*{}{}".format(
                colnum_string(i+3), fcff_row, colnum_string(i+3), cumulated_df_row)
            pv.append(current_fcff)
            #
            # current_fcff = 0
            # if fcff[i] is not None:
            #     current_fcff = fcff[i]
            # # fcff * df
            # pv.append(current_fcff * cumulated_df[i])

    def compute_terminals(self, d):
        roll_number = 14

        def add_rollng_number():
            nonlocal roll_number
            roll_number += 1
            return roll_number

        d.set('Terminal cash flow', d.get('FCFF').last(), add_rollng_number())
        d.set('Terminal cost of capital', d.get('Cost of capital').last(), add_rollng_number(), style='Percent')
        d.set('Terminal value',
              "{}/({}-{})".format(
                  d.get('Terminal cash flow').value(),
                  d.get('Terminal cost of capital').value(),
                  d.get('Revenue growth rate').last()),
              add_rollng_number())
        d.set('PV (Terminal value)',
              "{}*{}".format(
                  d.get('Terminal value').value(), d.get('Cumulated discount factor').last2()),
              add_rollng_number())

        mark = 'PV (FCFF)'
        d.set('PV (Cash flow over next 10 years)', "SUM({start}{row}:{end}{row})".format(
           start=d.get(mark).start(), end=d.get(mark).end(), row=d.get(mark).j), add_rollng_number())

        d.set('Sum of PV', "{}+{}".format(
            d.get('PV (Terminal value)').value(), d.get('PV (Cash flow over next 10 years)').value()), add_rollng_number())

        d.set('Value of operating assets', "{}".format(
            d.get('Sum of PV').value()), add_rollng_number())

        debt = 0
        if self.debt[-1] is not None:
            debt = self.debt[-1]
        d.set('- Debt', "{}".format(debt), add_rollng_number())

        # TODO
        d.set('- Minority interest', 0, add_rollng_number())
        d.set('+ Cash', "{}".format(self.cash[-1]), add_rollng_number())

        non_op = 0
        if type(self.investments) is list:
            # type: List[float]
            if self.investments[-1] is not None:
                non_op = self.investments[-1]
            elif self.investments[-2] is not None:
                colour_print("Latest investment was not defined. Fallback to previous year", bcolors.WARNING)
                non_op = self.investments[-2]
        d.set('+ Non-operating assets', "{}".format(non_op), add_rollng_number())

        d.set('Value of equity', "{}-{}-{}+{}+{}".format(
            d.get('Value of operating assets').value(),
            d.get('- Debt').value(),
            d.get('- Minority interest').value(),
            d.get('+ Cash').value(),
            d.get('+ Non-operating assets').value(),
        ), add_rollng_number())

        d.set('Number of shares', self.shares[-1], add_rollng_number())
        d.set('Estimated value / share', "{}/{}".format(
            d.get('Value of equity').value(), d.get('Number of shares').value()), add_rollng_number())

        ticker = self.get_ticker()
        avg_price = (ticker.info['regularMarketDayLow'] + ticker.info['regularMarketDayHigh']) / 2.
        d.set('Price', avg_price, add_rollng_number())
        d.set('Price as % of value', "{}/{}".format(
            d.get('Price').value(), d.get('Estimated value / share').value()), add_rollng_number(), style='Percent')

    def get_ticker(self):
        if self.cached_ticker is None:
            tick_name = self.tick
            if re.match(r'\d{4}$', self.tick):
                # Regex to match currency denomination express by 4 digits such as 9618.HK
                suffix = self.dataset.get_currency_suffix(self.sticky_price)
                tick_name = tick_name + '.{}'.format(suffix)

            # Test the ticker by obtaining beta
            ticker = yf.Ticker(tick_name)
            if 'beta' not in ticker.info:
                assert len(self.head) > 0
                initial_query = ' '.join(self.head.split()[:-1])
                print("Waiting to query Yahoo Finance server with '{}'".format(initial_query))
                ticker = yf.Ticker(get_symbol(initial_query))
            self.cached_ticker = ticker
        else:
            ticker = self.cached_ticker
        return ticker

    def compute_return_invested_capital(self, d):
        ## Invested capital
        # TODO =IF(operating_lease="Yes",
        #     IF(rnd_expense_cap="Yes",
        #        book_value_equity + book_value_debt - cash + adjust_debt_outstanding + rnd_converter,
        #        book_value_equity + book_value_debt - cash + adjust_debt_outstanding),
        #     IF(rnd_expense_cap="Yes",
        #        book_value_equity + book_value_debt - cash + rnd_converter,
        #        book_value_equity + book_value_debt - cash))
        # d['empty3'] = []
        # invested_capital = d['Invested Capital'] = []
        d.add_label('Return', 32)
        row = 33
        invested_capital = d.create_array('Invested Capital', row,)
        book_value_debt = 0
        if self.book_value_debt[-1] is not None:
            book_value_debt = self.book_value_debt[-1]
        current_ic = self.book_value_equity[-1] + book_value_debt - self.cash[-1]
        prev_ic = []
        # reinvestment = d['- Reinvestment']
        reinvestment = d.get('- Reinvestment')
        reinvestment_row = 8
        prev_ic.append(reinvestment.value())
        invested_capital.append(current_ic)
        for i in range(total_main_col-1):
            # prev_ic.append(current_ic)
            # if reinvestment[i] is not None:
            # current_ic += reinvestment[i]
            current_ic = '={}{}+{}{}'.format(
                colnum_string(i+2), row, colnum_string(i+3), reinvestment_row)
            invested_capital.append(current_ic)

        # invested_return = d['ROIC'] = []
        nopat_row = 7
        invested_return = d.create_array('ROIC', row+1, style='Percent')
        invested_return.append(None)
        for i in range(1, total_main_col):
            if (i+1) > 1:
                invested_return.append("={}{} / {}{}".format(
                    colnum_string(i+2), nopat_row,
                    colnum_string(i+1), row))
            # invested_return.append(d['NOPAT'][i] / prev_ic[i])

        # TODO end of cost of capital
        # invested_return.append(d['Cost of capital'][-1])
        # self.excel.wb.save('aaa.xlsx')
        # exit()


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


dcf = DCF('intc', country='United States')
# dcf = DCF('qcom', country='United States')
dcf.compute()
print("XXX", dcf)

# Damodaran main data page
# https://pages.stern.nyu.edu/~adamodar/New_Home_Page/datacurrent.html