#!/usr/bin/env python3

import sys
import pandas as pd
import csv
from openpyxl import Workbook


# Funções Anonimas Base
def df(arq_csv): return pd.read_csv(arq_csv)


def read_file(arq_xls): return pd.read_excel(arq_xls)


# Funções Anonimas
def pair_file(file1, file2, column): return pd.merge(df(file1), df(file2), left_on=column, right_on=column,
                                                how='inner', right_index=False, left_index=False)


def read_file(file): return df(file).head(50)


def convert_to_csv(file_xls, file_csv): return read_file(file_xls).to_csv(file_csv, encoding='utf-8', index=None,
                                                                           header=True)


def remove_duplicades_lines(file): return df(file).drop_duplicates()


def search(file, column, item): return print(df(file)[df(file)[column] == item])


help_program = ("Usage: python exceltool.py [OPTION] [FILE]\n" +
                "\t-h --help\t\thelp\n" +
                "\t-c --convert-to-csv\tconvert to csv\n" +
                "\t-cl --clean\t\tremove empty lines\n" +
                "\t-l --list\t\tlist csv\n" +
                "\t-p --pair\t\tpair two csv\n" +
                "\t-ce --convert-to-excel\tconvert to excel\n" +
                "\t-s --search\t\tsearch word\n" +
                "\t-col --column\t\tselect column" +
                "\t-w --word\t\tword"
                "Examples:\n" +
                "\texceltool.exe --pair file1.csv file2.csv --column column --output file\n" +
                "\texceltool.exe --convert-to-csv file.xls --output file\n" +
                "\texceltool.exe --convert-to-excel file.csv --output file\n" +
                "\texceltool.exe --clean file.csv --column column --output file\n" +
                "\texceltool.exe --file file.csv --column column --search word" +
                "\texceltool.exe --list file.csv")


# ultimo passo
def convert_to_excel(file_csv, file_xlsx):
    wb = Workbook()
    ws = wb.active
    with open(file_csv, 'r') as f:
        for row in csv.reader(f):
            ws.append(row)
        wb.save(file_xlsx + '.xlsx')


def clean_lines(list, column): return df(list).dropna(subset=[column], how='all')


if __name__ == "__main__":
    try:
        if sys.argv[1] == '-h' or sys.argv[1] == '--help':
            print(help_program)
        elif sys.argv[1] == '-c' and sys.argv[3] == '-o' or sys.argv[1] == '--convert-to-csv' and \
                sys.argv[3] == '--output':
            convert_to_csv(file_xls=str(sys.argv[2]), file_csv=str(sys.argv[4]))
        elif sys.argv[1] == '-cl' and sys.argv[3] == '-col' and sys.argv[5] == '-o' or sys.argv[1] == '--clean' and \
                sys.argv[3] == '--column' and sys.argv[5] == '--output':
            clean_lines(list=sys.argv[2], column=sys.argv[4]).to_csv(sys.argv[6] + '.csv', index=False)
        elif sys.argv[1] == '-l' or sys.argv[1] == '--list':
            print(read_file(str(sys.argv[2])))
        elif sys.argv[1] == '-p' and sys.argv[4] == '-col' and sys.argv[6] == '-o' or sys.argv[1] == '--pair' and \
            sys.argv[4] == '--column' and sys.argv[6] == '--output':
            pair_file(file1=sys.argv[2], file2=sys.argv[3], column=sys.argv[5]).to_csv(sys.argv[7] + '.csv', index=False)
        elif sys.argv[1] == '-ce' and sys.argv[3] == '-o' or sys.argv[1] == '--convert-to-excel' and \
                sys.argv[3] == '--output':
            convert_to_excel(arq_csv=str(sys.argv[2]), arq_xlsx=sys.argv[4])
        elif sys.argv[1] == '-f' and sys.argv[3] == '-col' and sys.argv[5] == '-s' or sys.argv[1] == '--file' and \
                sys.argv[3] == '--column' and sys.argv[5] == '--search':
            search(file=sys.argv[2], column=sys.argv[4], item=sys.argv[6])
        else:
            print(help_program)
    except IndexError as erro:
        print(erro)
        sys.exit(1)
    finally:
        print("Goodbye!")
