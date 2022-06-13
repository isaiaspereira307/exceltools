# -*- coding: UTF-8 -*-

#!/usr/bin/env python3


import pandas as pd
import csv
from openpyxl import (
        Workbook, 
        load_workbook
)
import argparse
from openpyxl.styles import (
    Font,
    PatternFill,
)
import matplotlib.pyplot as plt
from pandas.plotting import table


# Funções Anonimas Base
def df(arq_csv: str): return pd.read_csv(arq_csv)


def read_file(arq_xls: str): return pd.read_excel(arq_xls)


# Funções Anonimas
def pair_file(file1: str, file2: str, column: str): 
    return pd.merge(
            df(file1), 
            df(file2), 
            left_on=column, 
            right_on=column,
            how='inner', 
            right_index=False, 
            left_index=False
    )


def read_file(file: str): return df(file).head(50)


def convert_file_to_csv(file_xls: str, file_csv: str): 
    return read_file(file_xls).to_csv(
            file_csv, 
            encoding='utf-8', 
            index=None,
            header=True
    )


def remove_duplicades_lines(file): return df(file).drop_duplicates()


def search(list_correct, list, column: str, word: str): 
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


# Para alterar os estilos de formatação das células
def alterar_formatacao(
    file_xls: str,
    font_size: str,
    font_name: str,
    linha: int,
    coluna: int,
    valor: str,
    negrito: bool
):
    wb = Workbook()
    sheet = wb.active

    sheet.cell(row = linha, column = coluna).value = valor
    if negrito == True:
        sheet.cell(row = linha, column = coluna).font = Font(size = font_size, name = font_name, bold=True)
    else:
        sheet.cell(row = linha, column = coluna).font = Font(size = font_size, name = font_name)

    wb.save('styles.xlsx')



def mudar_cores():
    # Carregar dados para variável
    wb = load_workbook('test.xlsx')
    # Escolhe active sheet
    ws = wb.active
    # Deleta primeira coluna, que é somente índice
    ws.delete_cols(1)
    # Cabeçalho em negrito e fundo azul
    # Fill parameters
    my_fill = PatternFill(start_color='5399FF',
                    end_color='5399FF',
                    fill_type='solid')
    # Bold Parameter
    my_font = Font(bold=True)
    # Formata o cabeçalho
    my_header = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1']
    for cell in my_header:
        ws[cell].fill = my_fill
        ws[cell].font = my_font
        # Adiciona fórmula SUM
    ws['F1'] = 'Total'
    for i in range(2,22):
        ws['F' + str(i)] = f'=SUM(C{i}:E{i})'
        ws['F' + str(i)].font = my_font
        ws['F' + str(i)].fill = my_fill
        # Salva o arquivo
    wb.save('test.xlsx')


def salve_in_image(image: str):
    ax = plt.subplot(111, frame_on=False) # no visible frame
    ax.xaxis.set_visible(False)  # hide the x axis
    ax.yaxis.set_visible(False)  # hide the y axis

    table(ax, df)  # where df is your data frame

    plt.savefig(image)


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
parser.add_argument('-size','--font-size','-fonte_size', type=str, help='font size')
parser.add_argument('-nf','--name-font','-name_font', type=str, help='font name')
parser.add_argument('-B','--bold','-bold', type=str, help='Bold true or false')


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
            
    except UnicodeDecodeError as erro:
        print(erro)

    except IndexError as erro:
        print(erro)
        
