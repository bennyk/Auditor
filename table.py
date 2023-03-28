
from bs4 import BeautifulSoup
import bs4.element
from typing import Optional, List, Dict, Generator, Callable, Tuple
import re

from collections import namedtuple
from uuid import uuid1
import bootstrap


DivParcel = namedtuple('DivParcel',
                       'content_width content_height eff_height upper_margin height_margin line_height'
                       ' multiline_factor contents text')


def fix_currency(value):
    if value in ('-', '–'):
        return "NA"

    val = re.sub(r',', '', value)

    # optional close parenthesis
    m = re.match(r'\(\s*([.\d]+)\s*\)?', val)
    if m is not None:
        val = "-" + m.group(1)

    if re.search(r'\.', val):
        return float(val)
    return int(val)


class Dim:
    def __init__(self, rsc_val: str):
        val: str
        unit: str

        m = re.match(r'(\d+)(\w+)$', rsc_val)
        assert m is not None
        assert m.group(1).isnumeric() is True
        val = int(m.group(1))
        unit = m.group(2)
        self.val = val
        self.unit = unit

    def __gt__(self, other):
        if self.val > other.val:
            return True
        return False


class DimList:
    def __init__(self, rsc_val: str):
        self.list = []
        for v in rsc_val.split():
            self.list.append(Dim(v))


class Margins:
    # https://www.w3schools.com/css/css_margin.asp
    def __init__(self, rsc_val: str):
        self.top, self.right, self.bottom, self.left = [Dim(v) for v in rsc_val.split()]


class Props:
    # https://www.w3schools.com/cssref/pr_border-bottom.asp
    def __init__(self, val):
        self.val = val

    def thickness(self):
        m = re.search(r'(\d+)px', self.val)
        if m is not None:
            return int(m.group(1))
        return 0


class SuppressLineBlock:
    def __init__(self):
        self.out = bootstrap.Out
        self.sup_line_feed = True
        self.data = []

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if not self.sup_line_feed:
            print("\n", end='')
            self.out.write('\n')

    def write(self, index, msg):
        self.out.write("{}|{}|".format(index, msg))
        print("{}|{}|".format(index, msg), end='')
        self.sup_line_feed = False
        self.data.append(msg)


class Page:
    re_text = re.compile(r"([-–•]?[()/A-Za-z,&'’ ]+(?:[-–]\s*[A-Za-z][()/A-Za-z,&'’ ]+)*)", re.UNICODE)
    re_numerical = re.compile(r'([-–]|(?:\(\s*)?[0-9][0-9]*(?:,[0-9]{1,3})*(?:(\.[0-9]+))?(?:\s*\))?)', re.UNICODE)

    def __init__(self, page: bs4.element.Tag):
        self.page = page
        self.data = []

    def parse_top(self):
        for x in self.page.contents:
            if type(x) == bs4.element.NavigableString:
                continue

            elif type(x) == bs4.element.Tag:
                x: bs4.element.Tag
                if x.name == 'table':
                    self.parse_tbody(x)
                elif x.name == 'tbody':
                    self.parse_tr(x)
                elif x.name == 'tr':
                    self.parse_td(x)
                elif x.name == 'div':
                    self.parse_div(x)
                else:
                    pass
                    # assert False
            else:
                pass

    def parse_div(self, tab):
        for x in tab.contents:
            if type(x) == bs4.element.NavigableString:
                continue

            elif type(x) == bs4.element.Tag:
                x: bs4.element.Tag
                if x.name == 'table':
                    self.parse_tbody(x)

    def parse_tbody(self, tr):
        for x in tr.contents:
            if type(x) == bs4.element.NavigableString:
                continue

            elif type(x) == bs4.element.Tag:
                x: bs4.element.Tag
                if x.name == 'tbody':
                    self.parse_tr(x)

    def parse_tr(self, tr):
        for x in tr.contents:
            if type(x) == bs4.element.NavigableString:
                continue

            elif type(x) == bs4.element.Tag:
                x: bs4.element.Tag
                if x.name == 'tr':
                    self.parse_td(x)

    def parse_td(self, td):
        index = 0
        with SuppressLineBlock() as sup_line_block:
            for x in td.contents:
                if type(x) == bs4.element.NavigableString:
                    continue

                elif type(x) == bs4.element.Tag:
                    x: bs4.element.Tag

                    if x.name == 'td':
                        spanned = self.parse_span(x)
                        if len(spanned) == 0:
                            text = x.text.strip()
                            text = re.sub(r'\n', ' ', text)

                            if re.match(Page.re_numerical, text):
                                try:
                                    sup_line_block.write(index, fix_currency(text))
                                except ValueError:
                                    sup_line_block.write(index, text)

                            elif re.match(Page.re_text, text):
                                sup_line_block.write(index, text)

                            index += 1
                        else:
                            for a in spanned:
                                sup_line_block.write(index, a)

        # print(sup_line_block.data)
        self.data.append(sup_line_block.data)
        pass

    def parse_span(self, span):
        result = []
        for x in span.contents:
            if type(x) == bs4.element.NavigableString:
                continue

            elif type(x) == bs4.element.Tag:
                x: bs4.element.Tag
                if x.name == 'span':
                    result = self.parse_span(x)

                elif x.name == 'div':
                    pass

                elif x.name == 'img':
                    pass

                else:
                    # elif x.name == 'a':
                    #     result.append(x.text)
                    # elif x.name == 'sup':
                    #     result.append(x.text)
                    result.append(x.text)

        return result


class StyleParser:
    def __init__(self, gen: Generator):
        self.generator = gen
        self.hash = {}
        self.stack = []     # List[str]

    def compile_class(self, cls_name: str) -> Optional[Dict[str, str]]:
        if cls_name[0] != ".":
            cls_name = "." + cls_name

        if (cls_name, 'str') in self.hash:
            return self._compile(cls_name)

        result = None
        while True:
            sp, block = next(self.generator)

            sp1 = sp.span(1)
            cls_current = block[sp1[0]:sp1[1]]

            sp2 = sp.span(2)
            self.hash[cls_current, 'str'] = block[sp2[0]:sp2[1]]

            if cls_name == cls_current and (cls_current, 'str') in self.hash:
                result = self.compile_class(cls_name)
                break
        return result

    def push_null(self):
        self.stack.append('null')

    def push_id_selector(self, id_selector: str):
        self.stack.append('#' + id_selector)
        id_selector_cxt = ' '.join(self.stack[0:])
        if (id_selector_cxt, 'str') in self.hash:
            return self._compile(id_selector_cxt)

        while True:
            sp, block = next(self.generator)

            sp1 = sp.span(1)
            _ = block[sp1[0]:sp1[1]]  # type: str
            current_cxt = re.sub(r'\s+', ' ', _)

            sp2 = sp.span(2)
            self.hash[current_cxt, 'str'] = block[sp2[0]:sp2[1]]

            if (id_selector_cxt, 'str') in self.hash:
                result = self._compile(id_selector_cxt)
                break
        return result

    def push_inplace(self, style_str: str) -> Dict[str, str]:
        key = 'inplace-{}'.format(str(uuid1()))
        self.stack.append('#' + key)
        self.hash[key, 'str'] = style_str
        return self.hash[key, 'str']

    def pop_id_selector(self):
        self.stack.pop()

    def _compile(self, cls_name) -> Dict[str, str]:
        if (cls_name, 'compiled') not in self.hash:
            self.hash[cls_name, 'compiled'] = self.compile_string(self.hash[cls_name, 'str'])
        else:
            # print("memoized", cls_name)
            pass
        return self.hash[cls_name, 'compiled']

    def compile_string(self, string: str) -> Dict[str, str]:
        result = {}
        for x in string.split(';'):
            if x != '':
                m = re.match(r'([-–\w]+)\s*:\s*(.+)$', x)
                if m is not None:
                    # print(m.group(1), m.group(2))
                    result[m.group(1)] = m.group(2)
        # print(result)
        return result
