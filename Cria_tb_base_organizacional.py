# -*- coding: utf-8 -*-
"""
Created on Tue Dec 26 09:59:46 2023

@author: wendel.almeida
"""

import snowflake.connector
import pandas as pd
from tkinter import Tk, filedialog
import sys

def dt_bd():
    # Configurar a janela Tkinter
    root = Tk()
    
    try:
        # Abrir a caixa de diálogo para seleção do arquivo
        caminho_bd = filedialog.askopenfilename(title="Selecione o arquivo", filetypes=[("Arquivos Excel", "*.xlsx;*.xls"), ("Arquivos CSV", "*.csv")])

        # Fechar a janela Tkinter após a seleção do arquivo
        root.destroy()
        
        # Obter a extensão do arquivo
        extensao = caminho_bd.split('.')[-1].lower()

        # Verificar o tipo de arquivo selecionado e ler conforme necessário
        if extensao in ['xls', 'xlsx']:
            # Ler o arquivo Excel
            bd = pd.read_excel(caminho_bd)
        elif extensao == 'csv':
            # Tentar ler o arquivo CSV com tratamento de erro
            try:
                bd = pd.read_csv(caminho_bd,sep=';',encoding='latin1', on_bad_lines='skip')
            except pd.errors.ParserError as pe:
                print(f"Erro de parsing: {pe}")
                print("Pulando linhas com erro e tentando novamente...")
                # Tentar novamente pulando linhas com erro
                bd = pd.read_csv(caminho_bd, encoding='latin1', on_bad_lines='skip', engine='python')
        else:
            raise ValueError("Formato de arquivo não suportado.")

        # Exibir os dados
        print("Lendo os dados do arquivo:")

        return bd

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None

# Chamar a função
bd = dt_bd()

# Substituir valores NaN por um valor padrão (por exemplo, 0)
bd = bd.fillna(0)


# Se o DataFrame foi lido com sucesso, converta para UTF-8 e salve como CSV
##if bd is not None:
##   try:
##       bd.to_csv('bd_csv.csv', index=False, encoding='utf-8')
##       print("Arquivo convertido e salvo como CSV (UTF-8).")
##   except Exception as e:
##        print(f"Erro ao salvar o arquivo como CSV: {e}")


bd = bd.rename(columns = {
    'CÓDIGO EMPRESA':'COD_EMPRESA',
    'RAZÃO SOCIAL':'DESC_RAZAO_SOCIAL',
    'CÓDIGO UNIDADE':'COD_UNIDADE',
    'UNIDADE ORGANIZACIONAL':'DESC_UNIDADE',
    'UNIDADE ORGANIZACIONAL - RESUMIDA':'DESC_UNIDADE_RESUMIDA',
    'DATA INÍCIO':'DATA_INICIO',
    'STATUS':'STATUS',
    'TIPO':'TIPO',
    'CÓDIGO RELATORIO':'COD_RELATORIO',
    'PERMITE CADASTRAR':'PERMITE_CADASTRAR',
    'CÓDIGO NÍVEL':'COD_NIVEL',
    'NÍVEL':'DESC_NIVEL',
    'CÓDIGO NÍVEL SUPERIOR':'COD_NIVEL_SUPERIOR',
    'NÍVEL SUPERIOR':'DESC_NIVEL_SUPERIOR',
    'IDENTIFICAR POR ':'IDENTIFICAR_POR', 
    'GESTOR':'GESTOR'
    })

try:
    conn = snowflake.connector.connect(
        user='####',
        account='####',
        warehouse='####',
        database='####',
        schema='####',
        authenticator='externalbrowser'
    )
    
    print('Conectado ao banco de dados!')
    


    cursor = conn.cursor()

    nome_tabela = 'TB_BASE_ORGANIZACIONAL'

    delete_query = f'DELETE FROM GENTE_GESTAO.DB_INFORH.{nome_tabela}'
    cursor.execute(delete_query)

    colunas = ', '.join(bd.columns)


     # Construir a string de marcadores de posição para os valores
    marcadores_posicao = ', '.join(['%s' for _ in range(len(bd.columns))])


    # Inserir dados na tabela
    sql_query = f"INSERT INTO {nome_tabela} ({colunas}) VALUES ({marcadores_posicao})"
    values_list = [tuple(map(str, row)) for row in bd.values.tolist()]  # Converter todos os valores para string
    cursor.executemany(sql_query, values_list)

    # Commit para salvar as alterações
    conn.commit()





except snowflake.connector.errors.DatabaseError as e:
    print(f'Erro de conexão: {e}')

    # Encerrar a execução do script em caso de falha de conexão
    sys.exit()

finally:
    if conn:
        conn.close()    
        