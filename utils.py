import re
import yahooquery as yq


def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string


def excel_to_decimal(excel_column):
    decimal = 0
    for char in excel_column:
        decimal = decimal * 26 + ord(char.upper()) - 64
    return decimal


# TODO skipped first element in the list
def striped_average(l: [float], prefix):
    l = strip(l, prefix)
    return sum(l) / len(l)


def average(l: [float]):
    return sum(filter(None, l)) / len(l)


def strip(l, prefix, trim_last=False):
    if trim_last:
        return l[prefix:][:-1]
    return l[prefix:]


def strip2(l, prefix, trim_last=False):
    # TODO need to check against original strip implementation.
    if trim_last:
        return l[prefix:][:-1]

    _ = l[prefix:]
    if type(_[0]) is str:
        return list(map(lambda x: float(re.sub(r'(\d+)x', r'\1', x)), _))
    return _


def list_over_list(x, y, percent=False):
    if percent:
        # return list(map(lambda n1, n2: (n1 / n2), x, y))
        return list(map(lambda n1, n2: 0 if n1 is None else 100 * (n1 / n2), x, y))
    return list(map(lambda n1, n2: 0 if n1 is None else n1 / n2, x, y))


def list_multiply_list(x, y):
    return list(map(lambda n1, n2: n1*n2, x, y))


def list_add_list(x, y):
    a = map(lambda n1: 0 if n1 is None else n1, x)
    b = map(lambda n2: 0 if n2 is None else n2, y)
    return list(map(lambda n1, n2: n1 + n2, a, b))


def list_minus_list(x, y):
    a = map(lambda n1: 0 if n1 is None else n1, x)
    b = map(lambda n2: 0 if n2 is None else n2, y)
    return list(map(lambda n1, n2: n1 - n2, a, b))


def list_abs(x):
    return list(map(lambda n1: 0 if n1 is None else abs(n1), x))


def list_one(x, l):
    return [1.] * l

def list_negate(x):
    return list(map(lambda n1: -n1, x))

def cagr(l: [float]) -> float:
    # standard CAGR formula:
    # https://corporatefinanceinstitute.com/resources/valuation/what-is-cagr/
    # Otherwise we can use piecewise function to workaround complex number
    # h(x) = -|x|^(1/n) for x < 0
    # h(x) = |x|^(1/n) for x >= 0
    # alternate interpretation
    #   p = (l[len(l)-1] + abs(l[0])) / abs(l[0])
    #   https://www.exceldemy.com/how-to-calculate-cagr-in-excel-with-negative-number/?utm_source=pocket_saves

    # preceding year
    y = -1
    ri = 1/len(l)
    if l[0] < 0:
        a = l[len(l)-1] / l[0]
        if l[len(l)-1] >= 0:
            p = - abs(a) ** ri + y
        else:
            p = a ** ri + y
    else:
        p = (l[len(l)-1] / l[0]) ** ri + y
        if type(p) is complex:
            # if complex try force negation with abs
            assert l[len(l)-1] < 0
            a = l[len(l)-1] / l[0]
            p = - abs(a) ** ri + y
    assert type(p) is not complex
    return p


def zsum(a):
    return sum(list(map(lambda x: x if x is not None else 0, a)))


def get_symbol(query, preferred_exchange=''):
    try:
        data = yq.search(query)
    except ValueError: # Will catch JSONDecodeError
        print(query)
    else:
        quotes = data['quotes']
        if len(quotes) == 0:
            return 'No Symbol Found'

        symbol = quotes[0]['symbol']
        for quote in quotes:
            if quote['exchange'] == preferred_exchange:
                symbol = quote['symbol']
                break
        return symbol


