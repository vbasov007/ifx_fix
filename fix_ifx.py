"""
Usage:
    fix_ifx FILE_IN (-i)
    fix_ifx FILE_IN FILE_OUT [--all_fixes] [--remove_empty_rows_on_top=HEADER_LEFT_VALUE] [--replace_nan_with_dash]
    [--replace_sub_sup] [--fix_multiline_cells] [--unit_of_measure=UNIT_DEF_STR]...

Arguments:
    FILE_IN             Original xlsx file downloaded from infineon.com
    FILE_OUT            Fixed xlsx file
    UNIT_DEF_STR        String to define value of measure in col. Format - 'col_num:unit'
    HEADER_LEFT_VALUE   -

Options:
    -i --info
    -f --remove_empty_rows_on_top=HEADER_LEFT_VALUE     Default value = 'Product'
    -n --replace_nan_with_dash
    -s --replace_sub_sup
    -d --remove_in_cell_duplicates
    -a --all_fixes
    -u --unit_of_measure=UNIT_DEF_STR


"""

from docopt import docopt
import pandas as pd
import re
from show_file_info import print_header_value_variation_stat


def main():
    args = docopt(__doc__)

    a_all_fixes = args['--all_fixes']

    a_info = args['--info']

    if args['--remove_empty_rows_on_top']:
        left_header = args['--remove_empty_rows_on_top']
    else:
        left_header = 'Product'

    a_unit_of_measure_list = args['--unit_of_measure']

    file_in = args['FILE_IN']
    file_out = args['FILE_OUT']

    try:
        df = pd.read_excel(file_in)
    except Exception as e:
        print(e)
        print("Can't read {0}".format(file_in))
        return

    if args['--replace_nan_with_dash'] or a_all_fixes:
        replace_nan_with_dash(df)

    df.astype(str)

    if args['--remove_empty_rows_on_top'] or a_all_fixes:
        remove_empty_rows_on_top(df, left_header)

    if args['--fix_multiline_cells'] or a_all_fixes:
        fix_multiline_cells(df)

    if args['--replace_sub_sup'] or a_all_fixes:
        replace_sub_and_sup(df)

    if len(a_unit_of_measure_list) > 0:

        header_list = df.columns.values.tolist()
        header_dict = table_headers_dict(df)

        for item in a_unit_of_measure_list:

            col_name, val_to_append = parse_long_argument_with_colon(item)
            if col_name and val_to_append:
                header = arg_to_header(col_name, header_dict, header_list )
                if header:
                    append_str_to_all_in_col(df, header, val_to_append)

    if a_info:
        remove_empty_rows_on_top(df, left_header)
        print_header_value_variation_stat(df)
        return

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
    df.update(df.applymap(fix_nan))


def fix_multiline_cells(df):
    def fix_multiline(orig_val):
        v = str(orig_val)
        s = v.split("\r\n")
        s = set([w.strip() for w in s])
        s = list(s)
        s.sort()
        return ", ".join(s)

    df.update(df.applymap(fix_multiline))



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


def replace_empty_cells_with_dash(df):
    def dash_if_empty(x):
        if pd.isna(x):
            return '-'
        else:
            return x
    df.update(df.applymap(dash_if_empty))


def remove_empty_rows_on_top(df, left_header):

    if df.columns.values.tolist()[0] == left_header:
        return True

    for i in range(len(df.index)):
        if df.loc[i].values[0] == left_header:
            df.columns = df.loc[i].values
            df.drop(df.index[range(i + 1)], inplace=True)
            return True

    return False


def append_str_to_all_in_col(df, col_name, appended_str):
    if col_name in df.columns.values.tolist():
        df[col_name] = df[col_name] + appended_str


def parse_long_argument_with_colon(arg):

    spl = arg.split(':')
    if len(spl) == 2:
        col_name = spl[0].strip().strip('"').strip("'")
        val = spl[1]
        return col_name, val
    else:
        return '', ''

def table_headers_dict(df):

    headers = df.columns.values.tolist()

    out_dict = dict()
    i = 1
    for h in headers:
        out_dict.update({str(i): str(h)})
        i += 1

    return out_dict


def arg_to_header(arg, header_dict, header_list):

    if arg in header_list:
        return arg
    elif arg in header_dict:
        return header_dict[arg]
    else:
        return None

if __name__ == '__main__':
    main()
