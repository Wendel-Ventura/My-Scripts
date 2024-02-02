# -*- coding: utf-8 -*-
"""
Created on Wed Jan  3 15:59:06 2024

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
        caminho_df = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])

        # Fechar a janela Tkinter após a seleção do arquivo
        root.destroy()
        
        # Ler o arquivo Excel
        df = pd.read_excel(caminho_df)

        # Exibir os dados
        print("Lendo os dados do arquivo:")
        print(df)

        return df

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None

# Chamar a função
df = dt_df()

df = df.rename(columns={
    'Diretoria - Aj':'DIRETORIA',
    'Nº Centro de Custo':'COD_CENTRO_CUSTO',
    'Estrutura Organizacional':'DESC_ESTRUTURA_ORGANIZACIONAL',
    'Sigla':'SIGLA_2',
    'SIGLAÁREA/LOJA':'SIGLA' ,
    'Diretoria':'DIRETORIA2',
    'Região':'REGIAO',
    'DIRETOR/HEAD':'HEAD',
    'GESTOR/ESPEC RESPONSÁVEL':'GESTOR',
    'UNIDADE RESUMIDA':'UNIDADE_RESUMIDA',
    'Estrutura':'ESTRUTURA',
    'DESCRIÇÃO':'DESCRICAO'
    })

colunas_selecionadas=['DIRETORIA','COD_CENTRO_CUSTO', 'DESC_ESTRUTURA_ORGANIZACIONAL',
    'SIGLA', 'REGIAO', 'HEAD', 'GESTOR','UNIDADE_RESUMIDA', 'ESTRUTURA', 'DESCRICAO'
    ]

df = df[colunas_selecionadas]


df['DIRETORIA'] = df['DIRETORIA'].astype(str)



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

    nome_tabela = 'TB_CENTRO_CUSTO_ESTRUTURA'
    
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











