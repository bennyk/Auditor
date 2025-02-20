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

        for row in range(1, balance_sheet.max_row + 1):
            self.extract_data(balance_sheet, ['Balance Sheet | TIKR.com',
                                              r'Total Debt',
                                              r'Total Assets',
                                              ], "Balance")
        pass

    def write_save(self):
        out_wb = Workbook()
        ws = out_wb.active
        ws.title = "Earnings summary"

        WADS = r'Weighted Average Diluted Shares Outstanding'
        ws.column_dimensions["A"].width = 25

        j = 2
        total_revenues_idx = 2
        net_income_idx = 3
        adj_net_income_idx = 4
        epu_idx = 5
        epu_sen_idx = 6
        market_cap_idx = 7
        price_close_idx = 8
        shares_outstanding_idx = 9
        per_idx = 10
        dps_idx = 11
        dps_sen_idx = 12
        div_yield_idx = 13

        ws.cell(row=total_revenues_idx, column=1).value = "Total revenues"
        ws.cell(row=net_income_idx, column=1).value = "Net income"
        ws.cell(row=adj_net_income_idx, column=1).value = "Adj. net income"
        ws.cell(row=epu_idx, column=1).value = "Adj. EPU"
        ws.cell(row=epu_sen_idx, column=1).value = "Adj. EPU (sen)"
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

            ws.cell(row=net_income_idx, column=j).value = data[r'Net Income'][i]
            ws.cell(row=net_income_idx, column=j).number_format = FORMAT_NUMBER_00

            ws.cell(row=adj_net_income_idx, column=j).value = data[r'EBT Excl. Unusual Items'][i]
            ws.cell(row=adj_net_income_idx, column=j).number_format = FORMAT_NUMBER_00

            if data[WADS][i] is not None:
                ws.cell(row=epu_idx, column=j).value = f"={colnum_string(j)}{adj_net_income_idx}/{colnum_string(j)}{shares_outstanding_idx}"
                ws.cell(row=epu_idx, column=j).number_format = '0.0000'

                ws.cell(row=epu_sen_idx, column=j).value = f"=100*{colnum_string(j)}{epu_idx}"
                ws.cell(row=epu_sen_idx, column=j).number_format = FORMAT_NUMBER_00

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

        j = 2
        debt_to_assets_idx = 16
        ws.cell(row=debt_to_assets_idx, column=1).value = "Debt to Assets %"
        data = self.data["Balance"]
        for i in range(len(data[r'Balance Sheet | TIKR.com'])):
            ws.cell(row=15, column=j).value = data[r'Balance Sheet | TIKR.com'][i]

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
