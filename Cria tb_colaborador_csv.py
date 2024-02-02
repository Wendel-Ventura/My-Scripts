# -*- coding: utf-8 -*-
"""
Created on Tue Jan  2 08:21:02 2024

@author: wendel.almeida
"""

import snowflake.connector
import pandas as pd
from tkinter import Tk, filedialog
import sys
import snowflake.connector.pandas_tools as scpt


def prad(x):
    if isinstance(x, str):
        x = x.encode('utf-8').decode('utf-8') 
        # x = f'"{x}"'
    return x






def dt_bd():
    # Configurar a janela Tkinter
    root = Tk()
    
    try:
        # Abrir a caixa de diálogo para seleção do arquivo
        caminho_df = filedialog.askopenfilename(title="Selecione o arquivo", filetypes=[("Arquivos Excel", "*.xlsx;*.xls"), ("Arquivos CSV", "*.csv")])

        # Fechar a janela Tkinter após a seleção do arquivo
        root.destroy()
        
        # Obter a extensão do arquivo
        extensao = caminho_df.split('.')[-1].lower()

        # Verificar o tipo de arquivo selecionado e ler conforme necessário
        if extensao in ['xls', 'xlsx']:
            # Ler o arquivo Excel
            df = pd.read_excel(caminho_df,encoding='utf-8')
        elif extensao == 'csv':
            # Tentar ler o arquivo CSV com tratamento de erro
            try:
                df = pd.read_csv(caminho_df,sep=';',encoding='latin1', on_bad_lines='skip',skiprows=1)
                df = df.replace('-','')
            except pd.errors.ParserError as pe:
                print(f"Erro de parsing: {pe}")
                print("Pulando linhas com erro e tentando novamente...")
                # Tentar novamente pulando linhas com erro
                df = pd.read_csv(caminho_df, encoding='latin1', on_bad_lines='skip', engine='python')
        else:
            raise ValueError("Formato de arquivo não suportado.")

        # Exibir os dados
        print("Lendo os dados do arquivo:")
        # df = df.applymap(lambda x: x.encode('utf-8').decode('utf-8') if isinstance(x, str) else x)
        df = df.applymap(prad)
        # df = df.applymap(lambda x:"test" if isinstance(x, str) else x)

        return df

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None
    
    # Chamar a função
df = dt_bd()


df = df.rename(columns={
    'Código da Empresa':'COD_EMPRESA',
    'Razão Social':'DESC_RAZAO_SOCIAL',
    'Matrícula':'MATRICULA',
    'Colaborador':'NOME',
    'Data Admissão':'DATA_ADMISSAO',
    'Data Nascimento':'DATA_NASCIMENTO',
    'Idade':'IDADE',
    'Sexo':'SEXO',
    'PCD':'PCD',
    'Grau de Instrução':'GRAU_INSTITUICAO',
    'Logradouro':'LOGRADOURO_CASA',
    'Número':'NUMERO_CASA',
    'Complemento':'COMPLEMENTO_CASA',
    'Bairro':'BAIRRO_CASA',
    'CEP':'CEP_CASA',
    'Município':'MUNICIPIO_CASA',
    'Estado':'ESTADO_CASA',
    'País':'PAIS',
    'Código Sindicato':'CODIGO_SINDICATO',
    'Sindicato':'DESC_SINDICATO',
    'CNPJ':'CNPJ',
    'CódigoEstabelecimento':'COD_ESTABELECIMENTO',
    'Estabelecimento':'DESC_ESTABELECIMENTO',
    'Código Lotação Tributária':'COD_LOTACAO_TRIBUTARIO',
    'Lotação Tributária':'DESC_LOTACAO_TRIBUTARIA',
    'Código da Unidade Organizacional':'COD_UNIDADE_ORGANIZACIONAL',
    'Unidade Organizacional':'DESC_UNIDADE_ORGANIZACIONAL',
    'Código do Centro de Custo':'COD_CENTRO_CUSTO',
    'Centro de Custo':'DESC_CENTRO_CUSTO',
    'Código do Órgão Responsável':'COD_ORGAO_RESPONSAVEL',
    'Órgão Responsável':'DESC_ORGAO_RESPONSAVEL',
    'Código da Natureza Profissional':'COD_NATUREZA_PROFISSIONAL',
    'Natureza Profissional':'DESC_NATUREZA_PROFISSIONAL',
    'Código Categoria':'COD_CATEGORIA',
    'Categoria':'DESC_CATEGORIA',
    'Código Vínculo':'COD_VINCULO',
    'Vínculo':'DESC_VINCULO',
    'Codigo Jornada':'COD_JORNADA',
    'Jornada':'DESC_JORNADA',
    'Código Posição':'COD_POSICAO',
    'Posição':'DESC_POSICAO',
    'Tipo de Salário':'TIPO_SALARIO',
    'Código do Cargo':'COD_CARGO',
    ' Cargo':'DESC_CARGO',
    'Salário Atual':'SALARIO_ATUAL',
    'Código do CBO':'COD_CBO',
    'CBO':'DESC_CBO',
    'Telefone':'TELEFONE',
    'Celular':'CELULAR',
    'E-mail':'E-MAIL',
    'CPF':'CPF',
    'PIS/PASEP':'PIS/PASEP',
    'RG - Expedidor  - UF':'RG',
    'RG - Data Emissão':'DATA_EMISSAO_RG',
    'CTPS - Serie - UF':'CTPS_SERIE_UF',
    'Situação':'SITUACAO',
    'Data de Rescisão':'DATA_RECISAO',
    'Motivo Desligamento':'MOTIVO_DESLIGAMENTO'
    })


print('Dados recebidos e tratados')

df['SALARIO_ATUAL'] = df['SALARIO_ATUAL'].replace(',','.')
df['MATRICULA'] = df['MATRICULA']

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

    nome_tabela = 'TB_COLABORADORES_CSV'
    
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



