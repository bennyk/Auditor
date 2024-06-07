from operator import add, sub, mul, truediv
from openpyxl import Workbook, load_workbook, worksheet
from openpyxl.styles import Font, Alignment
import re
from utils import *
from collections import OrderedDict

total_main_col = 12


class ExcelWriter:
    start_col = 2
    row_margin = 1

    def __init__(self, tick: str):
        cls = self.__class__
        self.tick = tick
        # self.styles = styles
        self.ft = Font(name='Calibri', size=11)
        self.wb = Workbook()

        # type: worksheet.Worksheet
        self.sheet = self.wb.active
        self.sheet.title = 'sheet 1'

        self.cell = self.sheet.cell(row=1, column=cls.start_col)
        self.start_row_index = cls.row_margin+1
        self.end_row_index = len(self.tick) + self.start_row_index+1

        self.dict = None
        self.init_sheet()

    def init_sheet(self):
        sheet = self.sheet
        cell = sheet.cell(row=1, column=1)
        cell.value = self.tick.upper()
        sheet.cell(row=1, column=2).value = 'Base year'
        for i in range(1, total_main_col-1):
            cell = sheet.cell(row=1, column=i+2)
            cell.value = i
            cell.alignment = Alignment(horizontal='center')
        sheet.cell(row=1, column=total_main_col+1).value = 'Terminal year'

    def create_dict(self):
        self.dict = ExcelDict(self)
        return self.dict


def excel_calc(wb, cell):
    sheet = wb.active
    string = ''
    if cell.value[0] == '=':
        # print("XXX", cell.value[0])
        line = cell.value[1:]
        for x in re.findall(r'\w+|\S', line):
            m = re.match(r'([A-Za-z]+)(\d+)', x)
            if m is not None:
                cell = sheet.cell(
                    row=int(m.group(2)),
                    column=excel_to_decimal(m.group(1)))
                string += ' ' + str(cell.value)
            else:
                string += ' ' + x
    return calculate(string)


class ExcelDict:
    def __init__(self, excel: ExcelWriter):
        self.excel = excel

    def create_array(self, key, row, style='Comma'):
        start_col = 1
        sheet = self.excel.sheet
        sheet.column_dimensions[colnum_string(start_col)].width = 32
        cell = sheet.cell(row=row, column=start_col)
        cell.alignment = Alignment(wrapText=True)
        # TODO create array?
        cell.value = key
        return ExcelArray(row, self.excel, style=style)

    def add_label(self, label, row):
        sheet = self.excel.sheet
        cell = sheet.cell(row=row, column=1)
        cell.value = label

    def get(self, key):
        empty_count = 0
        result = None
        sheet = self.excel.sheet
        for j in range(2, 99):
            cell = sheet.cell(row=j, column=1)
            if cell.value is None:
                # TODO empty row tolerance
                if empty_count > 1:
                    break
                empty_count += 1
            else:
                if key == cell.value:
                    result = ExcelArray(j, self.excel)
                    break
        return result
    # Skipping accessors: getitem, setitem, delitem, iter, len

    def set(self, key, val, row, style='Comma'):
        sheet = self.excel.sheet
        cell = sheet.cell(row=row, column=1)
        cell.value = key

        cell = sheet.cell(row=row, column=2)
        cell.value = '={}'.format(val)
        cell.style = style
        if cell.style == 'Percent':
            cell.number_format = '0.00%'
        else:
            cell.number_format = '#,0.00'
        cell.font = Font(name='Calibri', size=11)


class ExcelArray:
    def __init__(self, row: int, excel: ExcelWriter, style: str='Comma'):
        self.excel = excel
        self.style = style

        self.i = 2
        self.j = row

    def append(self, val):
        sheet = self.excel.sheet
        cell = sheet.cell(row=self.j, column=self.i)
        sheet.column_dimensions[colnum_string(self.i)].width = 11
        cell.alignment = Alignment(wrapText=True)
        if (type(val) is int or type(val) is float) and val == 0:
            # Suppress zero value to empty string.
            cell.value = ''
        else:
            cell.value = val
        if self.style == 'Comma':
            if val != 0:
                cell.style = self.style
                cell.number_format = '#,0.00'
        elif self.style == 'Percent':
            cell.style = self.style
            cell.number_format = '0.00%'
        elif self.style == 'Ratio2':
            cell.number_format = '0.00'
        elif self.style == 'Ratio':
            # cell.style = style
            cell.number_format = '0.0000'
        else:
            assert False
        cell.font = self.excel.ft
        self.i += 1

    # TODO Need to remove hardcode numbers 3, 12 and 13
    def last(self):
        return "{}{}".format(colnum_string(13), self.j)

    def last2(self):
        return "{}{}".format(colnum_string(12), self.j)

    def value(self):
        return "{}{}".format(colnum_string(self.i), self.j)

    def start(self):
        return colnum_string(3)

    def end(self):
        return colnum_string(12)


class Calculator(object):
    # https://gist.github.com/maxkibble/1f0b4de51576ae75356c6a61b7aa1544
    op = {'+': add, '-': sub, '*': mul, '/': truediv}

    def to_suffix(self, s):
        st = []
        ret = ''
        tokens = s.split()
        for tok in tokens:
            if tok in ['*', '/']:
                while st and st[-1] in ['*', '/']:
                    ret += st.pop() + ' '
                st.append(tok)
            elif tok in ['+', '-']:
                while st and st[-1] != '(':
                    ret += st.pop() + ' '
                st.append(tok)
            elif tok == '(':
                st.append(tok)
            elif tok == ')':
                while st[-1] != '(':
                    ret += st.pop() + ' '
                st.pop()
            else:
                ret += tok + ' '
        while st:
            ret += st.pop() + ' '
        return ret

    def eva(self, s):
        st = []
        tokens = s.split()
        for tok in tokens:
            if tok not in self.op:
                st.append(float(tok))
            else:
                n1 = st.pop()
                n2 = st.pop()
                st.append(self.op[tok](n2, n1))
        return st.pop()

    def evaluate(self, string):
        # print(self.to_suffix(string))
        return self.eva(self.to_suffix(string))


# Input should split each operator/number with space, can handle float, negative and brackets
# cal = Calculator()
# print(cal.evaluate('( 20.0 / 2 ) + ( -3 * ( 1 + 2 ) )'))
def calculate(string):
    calc = Calculator()
    return calc.evaluate(string)
