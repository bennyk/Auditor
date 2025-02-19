import re

from openpyxl import load_workbook, Workbook
from openpyxl.styles.numbers import FORMAT_NUMBER_00, FORMAT_PERCENTAGE_00
from utils import colnum_string

path = '../spreads'


class ExcelSheet:
    def __init__(self):
        self.data = {}
        pass

    def extract_data(self, ws, target_labels):
        max_row, max_column = ws.max_row + 1, ws.max_column + 1
        self.data = {label: [] for label in target_labels}

        for i in range(1, max_row):
            row_label = ws[f'A{i}'].value
            if row_label in target_labels:
                for col in range(2, max_column):
                    value = ws.cell(row=i, column=col).value
                    if row_label == 'Price Close' and value is not None:
                        # Data was extracted from TIKR terminal
                        value = float(re.sub(r'MYR\s+', '', value)) if isinstance(value, str) else value
                    self.data[row_label].append(value)
        pass

    def parse_statement(self, name):
        wb = load_workbook(f"{path}/{name}.xlsx")
        self.data = {}
        income_sheet = wb['Income']

        for row in range(1, income_sheet.max_row+1):
            self.extract_data(income_sheet, ['Income Statement | TIKR.com',
                                             r'Total Revenues',
                                             r'Net Income',
                                             r'EBT Excl. Unusual Items',
                                             r'Weighted Average Diluted Shares Outstanding',
                                             r'Market Cap',
                                             r'Price Close',
                                             r'Dividends per share',
                                             ])
        pass

    def write_save(self):
        out_wb = Workbook()
        ws = out_wb.active
        ws.title = "Earnings summary"

        WADS = r'Weighted Average Diluted Shares Outstanding'
        ws.column_dimensions["A"].width = 25

        j = 2
        adj_net_income = 4
        epu = 5
        price_close = 7
        shares_outstanding = 8
        dps = 10
        ws.cell(row=2, column=1).value = "Total revenues"
        ws.cell(row=3, column=1).value = "Net income"
        ws.cell(row=adj_net_income, column=1).value = "Adj. net income"
        ws.cell(row=epu, column=1).value = "Adj. EPU (sen)"
        ws.cell(row=6, column=1).value = "Market Cap"
        ws.cell(row=price_close, column=1).value = "Price Close"
        ws.cell(row=shares_outstanding, column=1).value = "Shares outstanding"
        ws.cell(row=9, column=1).value = "PER"
        ws.cell(row=dps, column=1).value = "Dividends per share (sen)"
        ws.cell(row=11, column=1).value = "Dividends yield %"
        # ws.cell(row=11, column=1).value = "Dividends payout rate %"
        for i in range(len(self.data["Income Statement | TIKR.com"])):
            ws.cell(row=1, column=j).value = self.data[r'Income Statement | TIKR.com'][i]

            ws.cell(row=2, column=j).value = self.data[r'Total Revenues'][i]
            ws.cell(row=2, column=j).number_format = FORMAT_NUMBER_00

            ws.cell(row=3, column=j).value = self.data[r'Net Income'][i]
            ws.cell(row=3, column=j).number_format = FORMAT_NUMBER_00

            ws.cell(row=adj_net_income, column=j).value = self.data[r'EBT Excl. Unusual Items'][i]
            ws.cell(row=adj_net_income, column=j).number_format = FORMAT_NUMBER_00

            if self.data[WADS][i] is not None:
                ws.cell(row=epu, column=j).value = f"=100 * {colnum_string(j)}{adj_net_income}/{colnum_string(j)}{shares_outstanding}"
                ws.cell(row=epu, column=j).number_format = FORMAT_NUMBER_00

            ws.cell(row=6, column=j).value = self.data[r'Market Cap'][i]
            ws.cell(row=6, column=j).number_format = FORMAT_NUMBER_00

            ws.cell(row=price_close, column=j).value = self.data[r'Price Close'][i]
            ws.cell(row=price_close, column=j).number_format = FORMAT_NUMBER_00

            ws.cell(row=shares_outstanding, column=j).value = self.data[WADS][i]
            ws.cell(row=shares_outstanding, column=j).number_format = FORMAT_NUMBER_00

            if self.data[WADS][i] is not None:
                ws.cell(row=9, column=j).value = f"=100*{colnum_string(j)}{price_close}/{colnum_string(j)}{epu}"
                ws.cell(row=9, column=j).number_format = FORMAT_NUMBER_00

            ws.cell(row=10, column=j).value = self.data[r'Dividends per share'][i]*100 if self.data[r'Dividends per share'][i] is not None else ''
            ws.cell(row=10, column=j).number_format = FORMAT_NUMBER_00

            if self.data[WADS][i] is not None:
                ws.cell(row=11, column=j).value = f"={colnum_string(j)}{dps}/{colnum_string(j)}{price_close}/100"
                ws.cell(row=11, column=j).number_format = FORMAT_PERCENTAGE_00
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
