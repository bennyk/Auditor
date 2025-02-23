import re

from openpyxl import load_workbook, Workbook
from openpyxl.styles.numbers import FORMAT_NUMBER_00, FORMAT_PERCENTAGE_00
from utils import colnum_string, list_over_list

path = '../spreads'


class ExcelSheet:
    def __init__(self):
        self.data = {}
        pass

    def extract_data(self, ws, target_labels, tag):
        max_row, max_column = ws.max_row + 1, ws.max_column + 1
        data = {label: [] for label in target_labels}

        for i in range(1, max_row):
            row_label = ws[f'A{i}'].value
            if row_label in target_labels:
                for col in range(2, max_column):
                    value = ws.cell(row=i, column=col).value
                    if row_label == 'Price Close' and value is not None:
                        # Data was extracted from TIKR terminal
                        value = float(re.sub(r'MYR\s+', '', value)) if isinstance(value, str) else value
                    data[row_label].append(value)
        self.data[tag] = data
        pass

    def parse_statement(self, name):
        wb = load_workbook(f"{path}/{name}.xlsx")
        income_sheet = wb['Income']
        balance_sheet = wb['Balance']
        cash_flow = wb['Cash']

        for row in range(1, income_sheet.max_row+1):
            self.extract_data(income_sheet, ['Income Statement | TIKR.com',
                                             r'Total Revenues',
                                             r'Net Income',
                                             r'EBT Excl. Unusual Items',
                                             r'Weighted Average Diluted Shares Outstanding',
                                             r'Market Cap',
                                             r'Price Close',
                                             r'Dividends per share',
                                             ], 'Income')

        for row in range(1, cash_flow.max_row+1):
            self.extract_data(cash_flow, ['Cash Flow Statement | TIKR.com',
                                          r'Net Income',
                                          r'Total Depreciation, Depletion & Amortization',
                                          r'Total Asset Writedown',
                                          r'Provision and Write-off of Bad Debts',
                                          r'Acquisition of Real Estate Assets',
                                          ], 'Cash')

        for row in range(1, balance_sheet.max_row + 1):
            self.extract_data(balance_sheet, ['Balance Sheet | TIKR.com',
                                              r'Total Debt',
                                              r'Total Assets',
                                              ], "Balance")
        pass

    def parse_header_year(self, head):
        tikr_header = None
        if head == "Income":
            tikr_header = 'Income Statement | TIKR.com'
        elif head == "Balance":
            tikr_header = 'Balance Sheet | TIKR.com'
        elif head == "Cash":
            tikr_header = 'Cash Flow Statement | TIKR.com'
        else:
            raise "Invalid header"

        m = re.search(r'(\d{2})$', self.data[head][tikr_header][0])
        assert m is not None
        return int(m.group(1))

    def write_save(self):
        out_wb = Workbook()
        ws = out_wb.active
        ws.title = "Earnings summary"

        WADS = r'Weighted Average Diluted Shares Outstanding'
        ws.column_dimensions["A"].width = 25

        j = 2
        total_revenues_idx = 2
        revenues_growth_idx = 3
        net_income_idx = 4
        adj_net_income_idx = 5
        epu_idx = 6
        epu_sen_idx = 7
        epu_sen_growth_idx = 8
        market_cap_idx = 9
        price_close_idx = 10
        shares_outstanding_idx = 11
        per_idx = 12
        dps_idx = 13
        dps_sen_idx = 14
        div_yield_idx = 15

        ws.cell(row=1, column=1).value = "Income items / end of year"
        ws.cell(row=total_revenues_idx, column=1).value = "Total sales"
        ws.cell(row=revenues_growth_idx, column=1).value = "  Sales growth %"
        ws.cell(row=net_income_idx, column=1).value = "Net income"
        ws.cell(row=adj_net_income_idx, column=1).value = "Adj. net income"
        ws.cell(row=epu_idx, column=1).value = "Adj. EPU"
        ws.cell(row=epu_sen_idx, column=1).value = "Adj. EPU (sen)"
        ws.cell(row=epu_sen_growth_idx, column=1).value = "  EPU growth %"
        ws.cell(row=market_cap_idx, column=1).value = "Market Cap"
        ws.cell(row=price_close_idx, column=1).value = "Price Close"
        ws.cell(row=shares_outstanding_idx, column=1).value = "Shares outstanding"
        ws.cell(row=per_idx, column=1).value = "PER"
        ws.cell(row=dps_idx, column=1).value = "Dividends per share"
        ws.cell(row=dps_sen_idx, column=1).value = "Dividends per share (sen)"
        ws.cell(row=div_yield_idx, column=1).value = "Dividends yield %"
        # ws.cell(row=11, column=1).value = "Dividends payout rate %"

        data = self.data["Income"]
        for i in range(len(data[r'Income Statement | TIKR.com'])):
            ws.cell(row=1, column=j).value = data[r'Income Statement | TIKR.com'][i]

            ws.cell(row=total_revenues_idx, column=j).value = data[r'Total Revenues'][i]
            ws.cell(row=total_revenues_idx, column=j).number_format = FORMAT_NUMBER_00

            if i > 0 and data[r'Total Revenues'][i-1] is not None:
                ws.cell(row=revenues_growth_idx, column=j).value =\
                    f"={colnum_string(j)}{total_revenues_idx}/{colnum_string(j-1)}{total_revenues_idx}-1"
                ws.cell(row=revenues_growth_idx, column=j).number_format = FORMAT_PERCENTAGE_00

            ws.cell(row=net_income_idx, column=j).value = data[r'Net Income'][i]
            ws.cell(row=net_income_idx, column=j).number_format = FORMAT_NUMBER_00

            ws.cell(row=adj_net_income_idx, column=j).value = data[r'EBT Excl. Unusual Items'][i]
            ws.cell(row=adj_net_income_idx, column=j).number_format = FORMAT_NUMBER_00

            if data[WADS][i] is not None:
                ws.cell(row=epu_idx, column=j).value = f"={colnum_string(j)}{adj_net_income_idx}/{colnum_string(j)}{shares_outstanding_idx}"
                ws.cell(row=epu_idx, column=j).number_format = '0.0000'

                ws.cell(row=epu_sen_idx, column=j).value = f"=100*{colnum_string(j)}{epu_idx}"
                ws.cell(row=epu_sen_idx, column=j).number_format = FORMAT_NUMBER_00

                if i > 0 and data[WADS][i-1] is not None:
                    ws.cell(row=epu_sen_growth_idx, column=j).value = \
                        f"={colnum_string(j)}{epu_sen_idx}/{colnum_string(j-1)}{epu_sen_idx}-1"
                    ws.cell(row=epu_sen_growth_idx, column=j).number_format = FORMAT_PERCENTAGE_00

            ws.cell(row=market_cap_idx, column=j).value = data[r'Market Cap'][i]
            ws.cell(row=market_cap_idx, column=j).number_format = FORMAT_NUMBER_00

            ws.cell(row=price_close_idx, column=j).value = data[r'Price Close'][i]
            ws.cell(row=price_close_idx, column=j).number_format = FORMAT_NUMBER_00

            ws.cell(row=shares_outstanding_idx, column=j).value = data[WADS][i]
            ws.cell(row=shares_outstanding_idx, column=j).number_format = FORMAT_NUMBER_00

            if data[WADS][i] is not None:
                ws.cell(row=per_idx, column=j).value = f"={colnum_string(j)}{price_close_idx}/{colnum_string(j)}{epu_idx}"
                ws.cell(row=per_idx, column=j).number_format = FORMAT_NUMBER_00

            ws.cell(row=dps_idx, column=j).value = data[r'Dividends per share'][i] if data[r'Dividends per share'][i] is not None else ''
            ws.cell(row=dps_idx, column=j).number_format = '0.0000'

            ws.cell(row=dps_sen_idx, column=j).value = f"=100*{colnum_string(j)}{dps_idx}"
            ws.cell(row=dps_sen_idx, column=j).number_format = FORMAT_NUMBER_00

            if data[WADS][i] is not None:
                ws.cell(row=div_yield_idx, column=j).value = f"={colnum_string(j)}{dps_idx}/{colnum_string(j)}{price_close_idx}"
                ws.cell(row=div_yield_idx, column=j).number_format = FORMAT_PERCENTAGE_00
            j += 1

        # Income header year minus Cash header year. Years offset adjustment to IPO since pre listing.
        a = self.parse_header_year("Income")
        b = self.parse_header_year("Cash")
        j = 2 + b - a

        ffo_idx = 18
        ffo_per_share_idx = 19
        p_over_ffo_idx = 20
        affo_idx = 22
        affo_per_share_idx = 23
        p_over_affo_idx = 24
        ws.cell(row=17, column=1).value = "Cash items / end of year"
        ws.cell(row=ffo_idx, column=1).value = "FFO"
        ws.cell(row=ffo_per_share_idx, column=1).value = "FFO per share"
        ws.cell(row=p_over_ffo_idx, column=1).value = "P/FFO per share"

        ws.cell(row=affo_idx, column=1).value = "AFFO"
        ws.cell(row=affo_per_share_idx, column=1).value = "AFFO per share"
        ws.cell(row=p_over_affo_idx, column=1).value = "P/AFFO"

        data = self.data["Cash"]
        for i in range(len(data[r'Cash Flow Statement | TIKR.com'])):
            ws.cell(row=17, column=j).value = data[r'Cash Flow Statement | TIKR.com'][i]

            ffo = data['Net Income'][i]
            ffo += data['Total Depreciation, Depletion & Amortization'][i]
            if len(data['Total Asset Writedown']) > 0:
                if data['Total Asset Writedown'][i] is not None:
                    ffo += data['Total Asset Writedown'][i]

                if data['Provision and Write-off of Bad Debts'][i] is not None:
                    ffo += data['Provision and Write-off of Bad Debts'][i]

            ws.cell(row=ffo_idx, column=j).value = ffo
            ws.cell(row=ffo_idx, column=j).number_format = FORMAT_NUMBER_00

            ws.cell(row=ffo_per_share_idx, column=j).value = \
                f"={colnum_string(j)}{ffo_idx} / {colnum_string(j)}{shares_outstanding_idx}"
            ws.cell(row=ffo_per_share_idx, column=j).number_format = '0.0000'

            ws.cell(row=p_over_ffo_idx, column=j).value = f"={colnum_string(j)}{price_close_idx} / {colnum_string(j)}{ffo_per_share_idx}"
            ws.cell(row=p_over_ffo_idx, column=j).number_format = FORMAT_NUMBER_00

            affo = ffo
            affo += data['Acquisition of Real Estate Assets'][i]
            ws.cell(row=affo_idx, column=j).value = affo
            ws.cell(row=affo_idx, column=j).number_format = FORMAT_NUMBER_00

            ws.cell(row=affo_per_share_idx, column=j).value =\
                f"={colnum_string(j)}{affo_idx} / {colnum_string(j)}{shares_outstanding_idx}"
            ws.cell(row=affo_per_share_idx, column=j).number_format = '0.0000'

            ws.cell(row=p_over_affo_idx, column=j).value =\
                f"={colnum_string(j)}{price_close_idx} / {colnum_string(j)}{affo_per_share_idx}"
            ws.cell(row=p_over_affo_idx, column=j).number_format = FORMAT_NUMBER_00
            j += 1

        a = self.parse_header_year("Income")
        b = self.parse_header_year("Balance")
        j = 2 + b - a
        debt_to_assets_idx = 27
        ws.cell(row=26, column=1).value = "Balance items / end of year"
        ws.cell(row=debt_to_assets_idx, column=1).value = "Debt to Assets %"
        data = self.data["Balance"]
        for i in range(len(data[r'Balance Sheet | TIKR.com'])):
            ws.cell(row=26, column=j).value = data[r'Balance Sheet | TIKR.com'][i]

            ws.cell(row=debt_to_assets_idx, column=j).value = data["Total Debt"][i] / data["Total Assets"][i]
            ws.cell(row=debt_to_assets_idx, column=j).number_format = FORMAT_PERCENTAGE_00
            j += 1
        out_wb.save(f"xyz_report.xlsx")


def main():
    sheet = ExcelSheet()

    # for name in ['kipreit', 'igbreit', 'klcc', 'sunreit', 'axreit']:
    for name in ['kipreit']:
        sheet.parse_statement(name)
        sheet.write_save()


if __name__ == "__main__":
    main()


# KIPREIT TODOs
# 1. Add line revenue growth % and EPU growth %
# 2. P/FFO and P/AFFO
# 3. Pick up end-of-year items when tabulating shares outstanding (WADSO).
#    TIKR uses an average of outstanding shares instead, leading to a pessimistic/overvaluation when trading.
# 4. Does the calculated DPU (1.544 sen) not match the reported 1.66 sen?


