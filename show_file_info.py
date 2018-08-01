import pandas as pd
import re


def parameter_sort_key(string):

    r = re.match(r"^[-]?[\d]+\.?[\d]*", string)

    try:
        res = r.group(0)
        return float(res)
    except AttributeError:
        return 9999999


def is_sortable_as_number(string):

    r = re.match(r"^[-]?[\d]+\.?[\d]*", string)
    try:
        r.group(0)
        return True
    except AttributeError:
        return False


def sorted_smart(lst):
    as_number = [s for s in lst if is_sortable_as_number(s)]
    as_string = [s for s in lst if not is_sortable_as_number(s) and s != '-']

    as_number.sort(key=lambda x: parameter_sort_key(x))
    as_string.sort()

    if '-' in lst:
        as_string += '-'

    return as_number + as_string


def variations(df, col_header):
    a = set(df[col_header].tolist())
    lst = list(a)
    lst = [x if not pd.isna(x) else '-' for x in lst]
    res = list(map(str, lst))

    res = sorted_smart(res)

    return res


def print_limited_length(string, max_len=128):
    if len(string) < max_len:
        print(string[:max_len])
    else:
        print(string[:max_len], '...')


def print_header_value_variation_stat(df):

    print = print_limited_length

    headers = df.columns.values.tolist()
    i = 1

    max_len = max([len(h) for h in headers])

    for h in headers:
        v = variations(df, h)
        s1 = '{0}. "{1}" - {2}:'.format(i, h, len(v))
        s2 = " "*(max_len-len(s1)+20) + ', '.join(v)
        print(s1+s2)
        i += 1
