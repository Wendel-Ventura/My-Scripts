# -*- coding: utf-8 -*-
"""
Created on Tue Dec 26 13:33:51 2023

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

bd = bd.rename(columns={
     'Código Empresa':'COD_EMPRESA',
     'Empresa':'DESC_EMPRESA',
     'Código Estabelecimento':'COD_ESTABELECIMENTO',
     'Estabelecimento':'DESC_ESTABELECIMENTO',
     'Matrícula':'MATRICULA',
     'Colaborador':'NOME',
     'Admissão':'ADMISSAO',
     'Código Centro de Custo':'COD_CENTRO_CUSTO',
     'Centro de Custo':'DESC_CENTRO_CUSTO',
     'Código do Cargo':'COD_CARGO',
     'Cargo':'DESC_CARGO',
     'Início Afastamento':'INICIO_AFASTAMENTO',
     'Término Afastamento':'FIM_AFASTAMENTO',
     'Dias Afastamento':'DIAS_AFASTADOS',
     'Status Afastamento':'STATUS_AFASTAMENTO', 
     'Tipo Afastamento':'TIPO_AFASTAMENTO', 
     'Motivo Afastamento':'MOTIVO_AFASTAMENTO',
     'CID Afastamento':'CID_AFASTAMENTO',
     'Benefício de auxílio-doença concedido pela Previdência':'BENEFICIO_AUXILIO_DOENCA_PREVIDENCIA',
     'Início Atestado':'INICIO_ATESTADO',
     'Término Atestado':'TERMINO_ATESTADO', 
     'CID Atestado':'CID_ATESTADO', 
     'Dias Atestado':'DIAS_ATESTADOS'
     })
    
    
    
try:
    conn = snowflake.connector.connect(
        user='#####',
        account='#####',
        warehouse='#####',
        database='####',
        schema='#####',
        authenticator='externalbrowser'
    )

    print('Conectado ao banco de dados!')
    


    cursor = conn.cursor()

    nome_tabela = 'TB_AFASTADOS'

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

