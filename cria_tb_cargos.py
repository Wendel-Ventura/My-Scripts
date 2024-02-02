# -*- coding: utf-8 -*-
"""
Created on Tue Dec 26 07:59:00 2023

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
                bd = pd.read_csv(caminho_bd,sep=";" ,encoding='latin1', on_bad_lines='skip')
            except pd.errors.ParserError as pe:
                print(f"Erro de parsing: {pe}")
                print("Pulando linhas com erro e tentando novamente...")
                # Tentar novamente pulando linhas com erro
                bd = pd.read_csv(caminho_bd, encoding='latin1', on_bad_lines='skip', engine='python')
        else:
            raise ValueError("Formato de arquivo não suportado.")

        # Exibir os dados
        print("Lendo os dados do arquivo:")
        print(bd)

        return bd

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None

# Chamar a função
bd = dt_bd()

bd = bd.rename(columns={
    'Código Empresa':'COD_EMPRESA',
    'Razão Social'  : 'RAZAO_SOCIAL',
    'Código Cargo':'COD_CARGO',
    'Cargo':'DESC_CARGO',
    'Tipo Salário':'TIPO_SALARIO',
    'Status':'STATUS',
    'Data Início':'DATA_INICIO',
    'Permite Cadastrar Colaborador':'PERMITE_CADASTRAR_COLABORADOR',
    'Efetivo':'EFETIVO',
    'Código CBO':'COD_CBO',
    'CBO':'DESC_CBO',
    'Código Sindicato':'COD_SINDICATO',
    'Sindicato':'DESC_SINDICATO',
    'Código Nível Hierárquico':'COD_NIVEL_HIERARQUICO',
    'Nível Hierárquico':'DESC_NIVEL_HIERARQUICO',
    'Participa Cálculo Cota Aprendizagem':'PARTICIPA_CALCULO_PORCENTAGEM',
    'Cargo Público':'CARGO_PUBLICO',
    'Código Grupo de Cargos':'COD_GRUP_CARGOS',
    'Grupo de Cargos':'DESC_GRUPO_CARGOS'
    })

bd = bd.fillna(000)

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

    nome_tabela = 'TB_CARGOS'

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
        
        
