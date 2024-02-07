# -*- coding: utf-8 -*-
"""
Created on Tue Feb  6 16:19:09 2024

@author: wendel.almeida
"""
import snowflake.connector
import pandas as pd
from tkinter import Tk, filedialog
import sys
from datetime import datetime
import numpy as np
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

# Chamar a função para obter o DataFrame
df = dt_df()

# Lista de colunas na ordem desejada
Colunas_alteracao = [
    "Ano","Mês","Matrícula","Colaborador","Cargo","Salário","Data Admissão","Cód Estabelecimento","Centro de Custo","Sigla","Área/Loja",
    "Unidade Resumida","Diretoria","Porte","Regional","Região","Cód Unidade Organizacional","Unidade Organizacional","Código Posição","Posição",
    "Cód Cargo","Cód Nível Hierarq","Nível\nHierárquico","UO SUPERIOR","MATRÍCULA LÍDER","LÍDER","CARGO LÍDER","Sexo","Data de Nascimento","Idade",
    "Grau de Instrução","PCD","Situação","Data de Rescisão","Motivo Desligamento","Motivo Afastamento","CPF","RG","Grades Hay","Cód Praça","Praça",
    "Região2","Estrutura","Início Afastamento","Término Afastamento"]

# Renomear as colunas conforme necessário
df = df[Colunas_alteracao]

df = df.rename(columns={
    "Ano": "Ano",
    "Mês": "Mes",
    "Matrícula": "Matricula",
    "Colaborador": "Colaborador",
    "Cargo": "Cargo",
    "Salário": "Salario",
    "Data Admissão": "Data_Admissao",
    "Cód Estabelecimento": "Cod_Estabelecimento",
    "Centro de Custo": "Centro_de_Custo",
    "Sigla": "Sigla",
    "Área/Loja": "Area_Loja",
    "Unidade Resumida": "Unidade_Resumida",
    "Diretoria": "Diretoria",
    "Porte": "Porte",
    "Regional": "Regional",
    "Região": "Regiao",
    "Cód Unidade Organizacional": "Cod_Unidade_Organizacional",
    "Unidade Organizacional": "Unidade_Organizacional",
    "Código Posição": "Codigo_Posicao",
    "Posição": "Posicao",
    "Cód Cargo": "Cod_Cargo",
    "Cód Nível Hierarq": "Cod_Nivel_Hierarq",
    "Nível\nHierárquico": "Nivel_Hierarquico",
    "UO SUPERIOR": "UO_SUPERIOR",
    "MATRÍCULA LÍDER": "MATRICULA_LIDER",
    "LÍDER": "LIDER",
    "CARGO LÍDER": "CARGO_LIDER",
    "Sexo": "Sexo",
    "Data de Nascimento": "Data_de_Nascimento",
    "Idade": "Idade",
    "Grau de Instrução": "Grau_de_Instrucao",
    "PCD": "PCD",
    "Situação": "Situacao",
    "Data de Rescisão": "Data_de_Rescisao",
    "Motivo Desligamento": "Motivo_Desligamento",
    "Motivo Afastamento": "Motivo_Afastamento",
    "CPF": "CPF",
    "RG": "RG",
    "Grades Hay": "Grades_Hay",
    "Cód Praça": "Cod_Praca",
    "Praça": "Praca",
    "Região2": "Regiao2",
    "Estrutura": "Estrutura",
    "Início Afastamento": "Inicio_Afastamento",
    "Término Afastamento": "Termino_Afastamento"
})


# Limpeza de valores NaN
df.fillna('-', inplace=True)
# Limpeza de valores NaN
df.fillna('', inplace=True)
# Limpeza de valores NaN
df.fillna(0, inplace=True)

# Lista de colunas numéricas
colunas_numericas = ['Ano', 'Mes', 'Matricula', 'Salario', 'Cod_Estabelecimento', 'Centro_de_Custo',
                     'Cod_Cargo', 'Cod_Unidade_Organizacional', 'Cod_Praca', 'Cod_Nivel_Hierarq', 'MATRICULA_LIDER']

# Substituir valores vazios por NaN apenas nas colunas numéricas
df[colunas_numericas] = df[colunas_numericas].replace('', np.nan)

# Converta os valores NaN para zero (0) nas colunas numéricas
df[colunas_numericas] = df[colunas_numericas].fillna(0)



# Conversão de todos os campos para string
df = df.astype(str)

# Conversão para inteiros
cols_to_convert_int = ['Ano', 'Mes', 'Matricula', 'Cod_Unidade_Organizacional', 
                       'Cod_Estabelecimento', 'Centro_de_Custo', 'Cod_Cargo', 'Cod_Praca']

for col in cols_to_convert_int:
    df[col] = pd.to_numeric(df[col], errors='coerce', downcast='integer')

# Limpeza e conversão de Salario
df['Salario'] = df['Salario'].replace('nan', '')  # Substitui 'nan' por ''
df['Salario'] = pd.to_numeric(df['Salario'], errors='coerce', downcast='integer')

# Limpeza e conversão de Cod_Nivel_Hierarq
df['Cod_Nivel_Hierarq'] = df['Cod_Nivel_Hierarq'].replace('-', '')  # Substitui '-' por ''
df['Cod_Nivel_Hierarq'] = pd.to_numeric(df['Cod_Nivel_Hierarq'], errors='coerce', downcast='integer')

# Limpeza e conversão de MATRICULA_LIDER
df['MATRICULA_LIDER'] = df['MATRICULA_LIDER'].apply(lambda x: ''.join(filter(str.isdigit, str(x))))
df['MATRICULA_LIDER'].replace('', '0', inplace=True)  # Substitui strings vazias por '0'
df['MATRICULA_LIDER'] = df['MATRICULA_LIDER'].astype(int)  # Converte para inteiros



try:
    conn = snowflake.connector.connect(
        user='####',
        account='####',
        warehouse='#####',
        database='#####',
        schema='#####',
        authenticator='#####'
    )
    
    print('Conectado ao banco de dados!')
    
    cursor = conn.cursor()

    nome_tabela = 'TB_HIST_FOPAG'

    # Definir o tamanho do lote
    tamanho_lote = 1000

    # Dividir os dados em lotes menores
    num_rows = len(df)
    num_batches = num_rows // tamanho_lote + (1 if num_rows % tamanho_lote != 0 else 0)

    for i in range(num_batches):
        start_index = i * tamanho_lote
        end_index = min((i + 1) * tamanho_lote, num_rows)

        # Obter o lote atual
        batch_df = df.iloc[start_index:end_index]

        # Construir a string de marcadores de posição para os valores
        marcadores_posicao = ', '.join(['%s' for _ in range(len(batch_df.columns))])

        # Inserir dados na tabela
        sql_query = f"INSERT INTO {nome_tabela} ({', '.join(batch_df.columns)}) VALUES ({marcadores_posicao})"
        values_list = [tuple(row) for row in batch_df.values.tolist()]
        cursor.executemany(sql_query, values_list)

        # Commit para salvar as alterações
        conn.commit()
        print(f"Inseridos {len(batch_df)} registros do lote {i + 1}/{num_batches}")

    print('Todos os registros foram inseridos com sucesso!')

except snowflake.connector.errors.DatabaseError as e:
    print(f'Erro de conexão: {e}')

finally:
    if conn:
        conn.close()
        
