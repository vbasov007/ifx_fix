"""
Usage:
    fix_ifx FILE_IN (-i)
    fix_ifx FILE_IN FILE_OUT [--all_fixes] [--remove_first_rows=N_ROWS] [--replace_nan_with_dash] [--replace_sub_with_underscore]
    [--remove_in_cell_duplicates] [--value_designator=DES_STR]

Arguments:
    FILE_IN     Original xlsx file downloaded from infineon.com
    FILE_OUT    Fixed xlsx file

Options:
    -i --info
    -f --remove_first_rows=N_ROWS
    -n --replace_nan_with_dash
    -s --replace_sub_with_underscore
    -d --remove_in_cell_duplicates
    -a --all_fixes


"""

from docopt import docopt
import pandas as pd
import re

def main():
    args = docopt(__doc__)

    file_in = args['FILE_IN']
    file_out  = args['FILE_OUT']

    try:
        df = pd.read_excel(file_in)
    except Exception as e:
        print(e)
        print("Can't read {0}".format(file_in))
        return

    df = df.applymap(fixed)

    df.astype(str)

    replace_sub_and_sup(df)

    print(df.head())

    writer = pd.ExcelWriter(file_out, engine='xlsxwriter')
    df.to_excel(writer, index=False)
    writer.save()


def replace_sub_and_sup(df):

    header_names = df.columns.values.tolist()
    for name in header_names:
        fixed_name = re.sub('</sub>', '', re.sub('<sub>', '_', name))
        fixed_name = re.sub('<sup>2</sup>', chr(0x00b2), fixed_name)
        fixed_name = re.sub('<sup>3</sup>', chr(0x00b3), fixed_name)
        fixed_name = re.sub('<sup>', '', fixed_name)
        fixed_name = re.sub('</sup>', '', fixed_name)
        if fixed_name != name:
            df.rename(columns={name: fixed_name}, inplace=True)


def replace_nan_with_dash(df):
    def fix_nan(val):
        if pd.isna(val):
            return "-"
        else:
            return val
    df.applymap(fix_nan())


def remove_in_cell_duplicates(df):
    def fix_duplicate(orig_val):
        v = str(orig_val)
        s = v.split("\r\n")
        s = set([w.strip() for w in s])
        s = list(s)
        s.sort()
        return ", ".join(s)

    df.applymap(fix_duplicate())



def fixed(orig_val):

    if pd.isna(orig_val):
        return '-'

    v = str(orig_val)
    s = v.split("\r\n")
    s = set([w.strip() for w in s])
    s = list(s)
    if len(s) > 1:
        print('>>>>>>>>>>>>>>>>>>> {0}', s)
    s.sort()
    return ", ".join(s)

def remove_empty_first_row(df):
    pass


if __name__ == '__main__':
    main()
