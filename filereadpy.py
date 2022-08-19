# -*- coding: UTF-8 -*-

#!/usr/bin/env python


import pandas as pd
from hashlib import sha1
from katia import gerar_dair


def df(arq_csv: str):
    """Função que ler arquivo csv.""" 
    return pd.read_csv(arq_csv)


def pegar_ultimo_objeto(file_csv: str, option: str) -> int:
    """Pega o último objeto da lista."""
    return df(file_csv).iloc[-1][option]


def calcular_hash(file_csv: str) -> str:
    """Função que calcula hash SHA1."""
    with open(file_csv, 'r',encoding='utf-8') as content:
        f_content = content.read()
        f_hash: str = sha1(f_content.encode('utf-8')).hexdigest()
    return f_hash


def executar_automacao():
    """Executa a automação."""
    file_csv: str = ('automacao-dair.csv')
    while True:
        hash: str = calcular_hash(file_csv)
        if calcular_hash(file_csv) != hash:
            hash: str = calcular_hash(file_csv)
            id: int = pegar_ultimo_objeto(file_csv=file_csv, option='id')
            mes: int = pegar_ultimo_objeto(file_csv=file_csv, option='mes')
            ano: int = pegar_ultimo_objeto(file_csv=file_csv, option='ano')
            print(id, mes, ano)
            gerar_dair(id=id, mes=mes, ano=ano)


if __name__ == '__main__':
    try:
        executar_automacao()
    except UnicodeDecodeError as erro:
        print(erro)
    except IndexError as erro:
        print(erro)
    except KeyboardInterrupt as erro:
        print("Interrompido")