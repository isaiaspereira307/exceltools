#!/usr/bin/env python3

import sys
import numpy as np
import pandas as pd
import csv
from openpyxl import Workbook

# Funções Anonimas Base
data_frame = lambda arq_csv: pd.read_csv(arq_csv)
df1 = lambda arq_csv: pd.read_csv(arq_csv, usecols = ['Produto', 'Descricao', 'Quantidade', 'UM', 'Valor', 'ncm'])
df2 = lambda arq_csv: pd.read_csv(arq_csv, usecols = ['Produto Saida', 'Descricao Saida', 'Quantidade Saida', 'UM Saida', 'Valor Saida', 'ncm Saida'])
read_file = lambda arq_xls: pd.read_excel(arq_xls)

# Funções Anonimas 
parear_arquivo = lambda arq: pd.merge(df1(arq), df2(arq).rename(columns={'Produto Saida': 'Produto'}), how = 'outer', on = 'Produto')
listar_arquivo = lambda arq: data_frame(arq).head(50)
converter_para_csv = lambda arquivo_xls, arquivo_csv: read_file(arquivo_xls).to_csv(arquivo_csv, encoding='utf-8', index = None, header=True)

help = ("Usage: python exceltool.py [OPTION] [FILE]\n"+
		"\t-h --help\t\tajuda\n"+
		"\t-c --convert-to-csv\tconverter para csv\n"+
		"\t-cl --clean\t\texcluir linhas vazias\n"+
		"\t-l --list\t\tlistar csv\n"+
		"\t-p --pair\t\tparear dois csv\n"+
		"\t-ce --convert-to-excel\tconverter para excel")

#ultimo passo
def converter_para_excel(arquivo_csv):
	wb = Workbook()
	ws = wb.active
	with open(arquivo_csv, 'r') as f:
		for row in csv.reader(f):
			ws.append(row)
		wb.save('Listafinal.xlsx')

def limpar_linhas(lista_atual):
	df = pd.read_csv(lista_atual)
	df['Produto'].replace('', np.nan)
	df.dropna(subset=['Produto'], how='all')
	df.to_csv('Lista_definitiva.csv')

if __name__ == "__main__":
	try:
		if sys.argv[1] == '-h' or sys.argv[1] == '--help':
			print(help)
		elif sys.argv[1] == '-c' and sys.argv[3] == '-o' or sys.argv[1] == '--convert-to-csv' and sys.argv[3] == '--output':
			converter_para_csv(arquivo_xls=str(sys.argv[2]), arquivo_csv=str(sys.argv[4]))
		elif sys.argv[1] == '-cl' or sys.argv[1] == '--clean':
			limpar_linhas(lista_atual=sys.argv[2])
		elif sys.argv[1] == '-l' or sys.argv[1] == '--list':
			print(listar_arquivo(str(sys.argv[2])))
		elif sys.argv[1] == '-p' or sys.argv[1] == '--pair':
			parear_arquivo(sys.argv[2]).to_csv('listaatual.csv')
		elif sys.argv[1] == '-ce' or sys.argv[1] == '--convert-to-excel':
			converter_para_excel(arquivo_csv=str(sys.argv[2]))
		else:
			print(help)
	except IndexError:
		print(help)
		sys.exit(1)
