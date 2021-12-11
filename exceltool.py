#!/usr/bin/env python3

import sys
import pandas as pd
import csv
from openpyxl import Workbook


# Funções Anonimas Base
def df(arq_csv): return pd.read_csv(arq_csv)


def read_file(arq_xls): return pd.read_excel(arq_xls)


# Funções Anonimas
def parear_arquivo(arq1, arq2): return pd.merge(df(arq1), df(arq2), left_on='Produto', right_on='Produto',
                                                how='inner', right_index=False, left_index=False)


def listar_arquivo(arq): return df(arq).head(50)


def converter_para_csv(arq_xls, arq_csv): return read_file(arq_xls).to_csv(arq_csv, encoding='utf-8', index=None,
                                                                           header=True)


def remover_linhas_duplicadas(arq): return df(arq).drop_duplicates()


help_program = ("Usage: python exceltool.py [OPTION] [FILE]\n" +
                "\t-h --help\t\tajuda\n" +
                "\t-c --convert-to-csv\tconverter para csv\n" +
                "\t-cl --clean\t\texcluir linhas vazias\n" +
                "\t-l --list\t\tlistar csv\n" +
                "\t-p --pair\t\tparear dois csv\n" +
                "\t-ce --convert-to-excel\tconverter para excel\n" +
                "Examples:\n" +
                "\texceltool.exe --pair file1.csv file2.csv --output file\n" +
                "\texceltool.exe --convert-to-csv file.xls --output file\n" +
                "\texceltool.exe --convert-to-excel file.csv --output file\n" +
                "\texceltool.exe --clean file.csv --column column --output file\n" +
                "\texceltool.exe --list file.csv")


# ultimo passo
def converter_para_excel(arq_csv, arq_xlsx):
    wb = Workbook()
    ws = wb.active
    with open(arq_csv, 'r') as f:
        for row in csv.reader(f):
            ws.append(row)
        wb.save(arq_xlsx + '.xlsx')


def limpar_linhas(lista_atual, coluna): return df(lista_atual).dropna(subset=[coluna], how='all')


if __name__ == "__main__":
    try:
        if sys.argv[1] == '-h' or sys.argv[1] == '--help':
            print(help_program)
        elif sys.argv[1] == '-c' and sys.argv[3] == '-o' or sys.argv[1] == '--convert-to-csv' and \
                sys.argv[3] == '--output':
            converter_para_csv(arq_xls=str(sys.argv[2]), arq_csv=str(sys.argv[4]))
        elif sys.argv[1] == '-cl' and sys.argv[3] == '-col' and sys.argv[5] == '-o' or sys.argv[1] == '--clean' and \
                sys.argv[3] == '--column' and sys.argv[5] == '--output':
            limpar_linhas(lista_atual=sys.argv[2], coluna=sys.argv[4]).to_csv(sys.argv[6] + '.csv', index=False)
        elif sys.argv[1] == '-l' or sys.argv[1] == '--list':
            print(listar_arquivo(str(sys.argv[2])))
        elif sys.argv[1] == '-p' and sys.argv[4] == '-o' or sys.argv[1] == '--pair' and sys.argv[4] == '--output':
            parear_arquivo(sys.argv[2], sys.argv[3]).to_csv(sys.argv[5] + '.csv', index=False)
        elif sys.argv[1] == '-ce' and sys.argv[3] == '-o' or sys.argv[1] == '--convert-to-excel' and \
                sys.argv[3] == '--output':
            converter_para_excel(arq_csv=str(sys.argv[2]), arq_xlsx=sys.argv[4])
        else:
            print(help_program)
    except IndexError as erro:
        print(erro)
        sys.exit(1)
    finally:
        print("Goodbye!")
