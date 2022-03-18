#!/usr/bin/env python3

from ast import parse
from ensurepip import version
import sys
import pandas as pd
import csv
from openpyxl import Workbook
import argparse


# Funções Anonimas Base
def df(arq_csv): return pd.read_csv(arq_csv)


def read_file(arq_xls): return pd.read_excel(arq_xls)


# Funções Anonimas
def pair_file(file1, file2, column): return pd.merge(df(file1), df(file2), left_on=column, right_on=column,
                                                how='inner', right_index=False, left_index=False)


def read_file(file): return df(file).head(50)


def convert_file_to_csv(file_xls, file_csv): return read_file(file_xls).to_csv(file_csv, encoding='utf-8', index=None,
                                                                           header=True)


def remove_duplicades_lines(file): return df(file).drop_duplicates()


def search(list_correct, list, column, word): 
    for index, row in df(list_correct).interrows():
        print(row[column])
        print(index)
        df(list)[column].str.contains(word, regex=True)


def clean_lines(list, column): return df(list).dropna(subset=[column], how='all')


# ultimo passo
def convert_to_excel(file_csv, file_xlsx):
    wb = Workbook()
    ws = wb.active
    with open(file_csv, 'r') as f:
        for row in csv.reader(f):
            ws.append(row)
        wb.save(file_xlsx + '.xlsx')


parser = argparse.ArgumentParser(prog="exceltool", description='Excel Tool', 
                                    epilog="Author: Isaías Pereira",
                                    usage="%(prog)s [options]")

parser.version = "exceltool cli 1.0.0"
parser.add_argument('-v','--version', action="version")
parser.add_argument('-r','--read', type=str, help='read csv')
parser.add_argument('-c','--convert-to-csv', '-convert_to_csv', type=str, help='convert file to csv')
parser.add_argument('-cl','--clean', type=str, help='remove empty lines')
parser.add_argument('-p','--pair', nargs=2, type=str, help='pair two csv')
parser.add_argument('-ce','--convert-to-excel','-convert_to_excel', type=str, help='convert to excel')
parser.add_argument('-s','--search', type=str, help='search word')
parser.add_argument('-col','--column', type=str, help='select column')
parser.add_argument('-w','--word', type=str, help='word')
parser.add_argument('-o','--output', type=str, help='output file')

args = parser.parse_args()

if __name__=='__main__':
    try:
        if args.output:
            if args.pair:
                if args.column:
                    pair_file(file1=args.pair[0], file2=args.pair[1], column=args.column).to_csv(args.output + '.csv', index=False)
            elif args.convert_to_csv:
                convert_file_to_csv(file_xls=str(args.convert_to_csv), file_csv=str(args.output))
            elif args.convert_to_excel:
                convert_to_excel(arq_csv=str(args.convert_to_excel), arq_xlsx=args.output)
            elif args.clean:
                clean_lines(list=args.clean, column=args.column).to_csv(args.output + '.csv', index=False)
        elif args.read:
            print(read_file(str(args.read)))
        elif args.search:
            search(file=args.search, column=args.column, word=args.word)
            

    except IndexError as erro:
        print(erro)
        sys.exit(1)
        