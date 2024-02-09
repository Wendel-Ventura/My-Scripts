# -*- coding: utf-8 -*-
"""
Created on Thu Feb  8 12:09:46 2024
@author: wendel.almeida
"""
import snowflake.connector
import pandas as pd
from tkinter import Tk, filedialog
import numpy as np

############################ Recebendo os dados
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
                bd = pd.read_csv(caminho_bd, sep=';', encoding='latin1', on_bad_lines='skip', skiprows=(1))
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
df = dt_bd()

df = df.rename(columns={
    'Código da Empresa': 'CODIGO_DA_EMPRESA',
    'Razão Social': 'RAZAO_SOCIAL',
    'Matrícula eSocial': 'MATRICULA_ESOCIAL',
    'Matrícula': 'MATRICULA',
    'Colaborador': 'COLABORADOR',
    'Nome Social': 'NOME_SOCIAL',
    'ID Pessoa': 'ID_PESSOA',
    'Data Admissão': 'DATA_ADMISSAO',
    'Data de Vencimento ASO': 'DATA_DE_VENCIMENTO_ASO',
    'Data de Vencimento Contrato de Trabalho': 'DATA_DE_VENCIMENTO_CONTRATO_DE_TRABALHO',
    'Data de Vencimento Prorrogação de Contrato': 'DATA_DE_VENCIMENTO_PRORROGACAO_DE_CONTRATO',
    'Data Nascimento': 'DATA_NASCIMENTO',
    'Idade': 'IDADE',
    'Sexo': 'SEXO',
    'Gênero': 'GENERO',
    'Estado Civil': 'ESTADO_CIVIL',
    'Raça/Cor': 'RACA_COR',
    'País de Nacionalidade': 'PAIS_DE_NACIONALIDADE',
    'Tipo da Primeira Filiação': 'TIPO_DA_PRIMEIRA_FILIACAO',
    'Primeira Filiação': 'PRIMEIRA_FILIACAO',
    'Tipo da Segunda Filiação': 'TIPO_DA_SEGUNDA_FILIACAO',
    'Segunda Filiação': 'SEGUNDA_FILIACAO',
    'PCD': 'PCD',
    'Tipo de Deficiência': 'TIPO_DE_DEFICIENCIA',
    'Grau de Instrução': 'GRAU_DE_INSTRUCAO',
    'Tipo de Logradouro': 'TIPO_DE_LOGRADOURO',
    'Logradouro': 'LOGRADOURO',
    'Número': 'NUMERO',
    'Complemento': 'COMPLEMENTO',
    'Bairro': 'BAIRRO',
    'CEP': 'CEP',
    'Código do Município': 'CODIGO_DO_MUNICIPIO',
    'Município': 'MUNICIPIO',
    'Estado': 'ESTADO',
    'País': 'PAIS',
    'Código Sindicato': 'CODIGO_SINDICATO',
    'Sindicato': 'SINDICATO',
    'Primeiro Emprego': 'PRIMEIRO_EMPREGO',
    'CNPJ': 'CNPJ',
    'Código Estabelecimento': 'CODIGO_ESTABELECIMENTO',
    'Estabelecimento': 'ESTABELECIMENTO',
    'Código Lotação Tributária': 'CODIGO_LOTACAO_TRIBUTARIA',
    'Lotação Tributária': 'LOTACAO_TRIBUTARIA',
    'Código da Unidade Organizacional': 'CODIGO_DA_UNIDADE_ORGANIZACIONAL',
    'Unidade Organizacional': 'UNIDADE_ORGANIZACIONAL',
    'Código Nível Organizacional': 'CODIGO_NIVEL_ORGANIZACIONAL',
    'Nível Organizacional': 'NIVEL_ORGANIZACIONAL',
    'Código do Centro de Custo': 'CODIGO_DO_CENTRO_DE_CUSTO',
    'Centro de Custo': 'CENTRO_DE_CUSTO',
    'Código Área de Atuação': 'CODIGO_AREA_DE_ATUACAO',
    'Área de Atuação': 'AREA_DE_ATUACAO',
    'Código do Órgão Responsável': 'CODIGO_DO_ORGAO_RESPONSAVEL',
    'Órgão Responsável': 'ORGAO_RESPONSAVEL',
    'Código da Natureza Profissional': 'CODIGO_DA_NATUREZA_PROFISSIONAL',
    'Natureza Profissional': 'NATUREZA_PROFISSIONAL',
    'Código Categoria': 'CODIGO_CATEGORIA',
    'Categoria': 'CATEGORIA',
    'Código Vínculo': 'CODIGO_VINCULO',
    'Vínculo': 'VINCULO',
    'Codigo Jornada': 'CODIGO_JORNADA',
    'Jornada': 'JORNADA',
    'Carga Horária Mensal': 'CARGA_HORARIA_MENSAL',
    'Carga Horária Semanal': 'CARGA_HORARIA_SEMANAL',
    'Código Posição': 'CODIGO_POSICAO',
    'Posição': 'POSICAO',
    'Código do Cargo': 'CODIGO_DO_CARGO',
    'Cargo': 'CARGO',
    'Tipo de Salário': 'TIPO_DE_SALARIO',
    'Salário Atual': 'SALARIO_ATUAL',
    'Código da Função': 'CODIGO_DA_FUNCAO',
    'Função': 'FUNCAO',
    'Código do CBO': 'CODIGO_DO_CBO',
    'CBO': 'CBO',
    'Telefone': 'TELEFONE',
    'Celular': 'CELULAR',
    'E-mail Corporativo': 'E_MAIL_CORPORATIVO',
    'E-mail Particular': 'E_MAIL_PARTICULAR',
    'CPF/CNPJ': 'CPF_CNPJ',
    'PIS/PASEP': 'PIS_PASEP',
    'RG - Expedidor  - UF': 'RG_EXPEDIDOR_UF',
    'RG - Data Emissão': 'RG_DATA_EMISSAO',
    'CTPS - Serie - UF/ CTPS Digital': 'CTPS_SERIE_UF_CTPS_DIGITAL',
    'Tipo de Situação': 'TIPO_DE_SITUACAO',
    'Situação': 'SITUACAO',
    'Data de Rescisão': 'DATA_DE_RESCISAO',
    'Motivo Desligamento': 'MOTIVO_DESLIGAMENTO'
    })

# Convertendo todas as colunas para string
df = df.astype(str)

# Substituir valores vazios por NaN
df['SALARIO_ATUAL'] = df['SALARIO_ATUAL'].replace('', np.nan)


for col2 in ['SALARIO_ATUAL']:
    # Remover caracteres não numéricos antes da conversão
    df[col2] = df[col2].str.replace('[^\d,]', '', regex=True)
    # Converta para float, especificando o separador decimal como vírgula
    df[col2] = pd.to_numeric(df[col2].str.replace(',', '.'), errors='coerce').fillna(0)



# Definindo colunas que devem ser tratadas como numéricas
colunas_numericas = ['CARGA_HORARIA_MENSAL','CARGA_HORARIA_SEMANAL','CODIGO_DA_EMPRESA',
                     'MATRICULA_ESOCIAL','MATRICULA','ID_PESSOA','IDADE','CODIGO_DO_MUNICIPIO',
                     'CODIGO_SINDICATO','CODIGO_ESTABELECIMENTO','CODIGO_LOTACAO_TRIBUTARIA',
                     'CODIGO_NIVEL_ORGANIZACIONAL','CODIGO_DO_CENTRO_DE_CUSTO',
                     'CODIGO_AREA_DE_ATUACAO','CODIGO_DA_NATUREZA_PROFISSIONAL','CODIGO_CATEGORIA',
                     'CODIGO_VINCULO','CODIGO_POSICAO','CODIGO_DO_CARGO','CODIGO_DO_CBO','CODIGO_DA_FUNCAO',
                     'CODIGO_DO_ORGAO_RESPONSAVEL',]



# Função para converter valores para float, tratando valores vazios como 0
def convert_to_float(value):
    try:
        return float(value)
    except ValueError:
        return 0.0

# Aplicar a função de conversão para cada coluna numérica
for col in colunas_numericas:
    df[col] = df[col].apply(convert_to_float)

# Definindo colunas que devem ser tratadas como inteiras
colunas_inteiras = ['CARGA_HORARIA_MENSAL','CARGA_HORARIA_SEMANAL',
                    'CODIGO_DA_EMPRESA','MATRICULA_ESOCIAL',
                    'MATRICULA','ID_PESSOA','IDADE','CODIGO_DO_MUNICIPIO',
                    'CODIGO_SINDICATO','CODIGO_ESTABELECIMENTO','CODIGO_LOTACAO_TRIBUTARIA',
                    'CODIGO_NIVEL_ORGANIZACIONAL','CODIGO_DO_CENTRO_DE_CUSTO',
                    'CODIGO_AREA_DE_ATUACAO','CODIGO_DA_NATUREZA_PROFISSIONAL','CODIGO_CATEGORIA',
                    'CODIGO_VINCULO','CODIGO_POSICAO','CODIGO_DO_CARGO','CODIGO_DO_CBO',
                    'CODIGO_DA_FUNCAO','CODIGO_DO_ORGAO_RESPONSAVEL']

# Convertendo para tipo inteiro (downcast='integer')
for col in colunas_inteiras:
    df[col] = pd.to_numeric(df[col], errors='coerce', downcast='integer').fillna(0)

# Definindo colunas que devem ser tratadas como texto
colunas_txt = ['RAZAO_SOCIAL', 'COLABORADOR','NOME_SOCIAL', 'DATA_ADMISSAO',
               'DATA_DE_VENCIMENTO_ASO', 'DATA_DE_VENCIMENTO_CONTRATO_DE_TRABALHO',
               'DATA_DE_VENCIMENTO_PRORROGACAO_DE_CONTRATO', 'DATA_NASCIMENTO', 'SEXO',
               'GENERO', 'ESTADO_CIVIL', 'RACA_COR', 'PAIS_DE_NACIONALIDADE',
               'TIPO_DA_PRIMEIRA_FILIACAO', 'PRIMEIRA_FILIACAO', 'TIPO_DA_SEGUNDA_FILIACAO',
               'SEGUNDA_FILIACAO', 'PCD', 'TIPO_DE_DEFICIENCIA', 'GRAU_DE_INSTRUCAO',
               'TIPO_DE_LOGRADOURO', 'LOGRADOURO', 'NUMERO', 'COMPLEMENTO', 'BAIRRO',
               'CEP', 'MUNICIPIO', 'ESTADO', 'PAIS', 'SINDICATO', 'PRIMEIRO_EMPREGO',
               'CNPJ', 'ESTABELECIMENTO', 'LOTACAO_TRIBUTARIA', 'UNIDADE_ORGANIZACIONAL',
               'NIVEL_ORGANIZACIONAL', 'CENTRO_DE_CUSTO', 'AREA_DE_ATUACAO', 'ORGAO_RESPONSAVEL',
               'NATUREZA_PROFISSIONAL', 'CATEGORIA', 'VINCULO', 'JORNADA', 'POSICAO', 'CARGO',
               'TIPO_DE_SALARIO', 'FUNCAO', 'CBO', 'TELEFONE', 'CELULAR', 'E_MAIL_CORPORATIVO',
               'E_MAIL_PARTICULAR', 'CPF_CNPJ', 'PIS_PASEP', 'RG_EXPEDIDOR_UF', 'RG_DATA_EMISSAO',
               'CTPS_SERIE_UF_CTPS_DIGITAL', 'TIPO_DE_SITUACAO', 'SITUACAO', 'DATA_DE_RESCISAO',
               'MOTIVO_DESLIGAMENTO']

# Substituindo 'nan' por NaN em colunas de texto
for col in colunas_txt:
    df[col] = df[col].replace('nan', np.nan)

# Preenchendo valores NaN com '-' em colunas de texto
for col in colunas_txt:
    df[col] = df[col].fillna('-')



try:
    conn = snowflake.connector.connect(
        user='',
        account='',
        warehouse='',
        database='',
        schema='',
        authenticator='externalbrowser'
    )
        
    print('Conectado ao banco de dados!')
    
    cursor = conn.cursor()

    nome_tabela = 'DADOS_BRUTO_COLABORADORES_CSV'

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
    print('Todos os registros foram inseridos com sucesso!')

except snowflake.connector.errors.DatabaseError as e:
    print(f'Erro de conexão: {e}')

finally:
    if conn:
        conn.close()


