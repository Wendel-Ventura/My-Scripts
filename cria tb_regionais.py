# -*- coding: utf-8 -*-
"""
Created on Thu Jan  4 11:08:12 2024

@author: wendel.almeida
"""


import snowflake.connector
import pandas as pd
from tkinter import Tk, filedialog
import sys
import snowflake.connector.pandas_tools as scpt



############################ Recebendo os dados
def dt_df():
    # Configurar a janela Tkinter
    root = Tk()

    try:
        # Abrir a caixa de diálogo para seleção do arquivo
        caminho_df = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Arquivos Excel", "*.xlsx;*.xls;*.xlsm")])

        # Fechar a janela Tkinter após a seleção do arquivo
        root.destroy()

        # Verificar se o usuário cancelou a seleção
        if not caminho_df:
            print("Seleção de arquivo cancelada.")
            return None

        # Ler o arquivo Excel
        df = pd.read_excel(caminho_df, skiprows=2)

        # Exibir os dados
        print("Lendo os dados do arquivo:")
        print(df)

        return df

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None

# Chamar a função
df = dt_df()


colunas=['Loja', 'Região', 'Regional', 'Porte']

df = df[colunas]

df = df.rename(columns={
    'Loja':'LOJA',
    'Região':'REGIAO',
    'Regional':'REGIONAL',
    'Porte':'PORTE'
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

    nome_tabela = 'TB_REGIONAIS'
    
    print('deletando dados')
    delete_query = f'DELETE FROM GENTE_GESTAO.DB_INFORH.{nome_tabela}'
    cursor.execute(delete_query)
    print('dados deletados')
    print('Iniciando commit')

    colunas = ', '.join(df.columns)


     # Construir a string de marcadores de posição para os valores
    marcadores_posicao = ', '.join(['%s' for _ in range(len(df.columns))])


    # Inserir dados na tabela
    # sql_query = f"INSERT INTO {nome_tabela} ({colunas}) VALUES ({marcadores_posicao})"
    # values_list = [tuple(map(str, row)) for row in df.values.tolist()]  # Converter todos os valores para string
    # cursor.white_pandas(df=df)
    # cursor.executemany(sql_query, values_list)
    
    scpt.write_pandas(conn=conn,df=df,table_name=nome_tabela)

    # Commit para salvar as alterações
    conn.commit()

    print('Valores commitados com sucesso, Verificar a tabela')



except snowflake.connector.errors.DatabaseError as e:
    print(f'Erro de conexão: {e}')

    # Encerrar a execução do script em caso de falha de conexão
    sys.exit()

finally:
    if conn:
        conn.close()    



