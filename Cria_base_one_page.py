# -*- coding: utf-8 -*-
"""
Created on Thu Nov 23 08:52:21 2023

@author: wendel.almeida
"""

import pandas as pd
from tkinter import Tk, filedialog

############################ Recebendo os dados
def dt_bd():
    # Configurar a janela Tkinter
    root = Tk()
    
    try:
        # Abrir a caixa de diálogo para seleção do arquivo
        caminho_bd = filedialog.askopenfilename(title="Selecione o arquivo Excel", filetypes=[("Arquivos Excel", "*.xlsx;*.xls")])

        # Fechar a janela Tkinter após a seleção do arquivo
        root.destroy()
        
        # Ler o arquivo Excel
        bd = pd.read_excel(caminho_bd)

        # Exibir os dados
        print("Lendo os dados do arquivo:")
        print(bd)

        return bd

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None

# Chamar a função
bd = dt_bd()

############################ Selecionando e Manipulando os dados

colunas_selecionadas = ['Ano', 'Mês', 'Matrícula', 'Colaborador', 'Cargo', 'Salário', 'Data Admissão', 'Status final no mês', 'sigla - AJ',
                         'Status colaborador', 'Diretoria final', 'Unidade Resumida 24 com 23 (SEEEROO)', 'Cód Estabelecimento', 'Porte', 'Regional',
                         'Região', 'Código Posição', 'Posição', 'Cód Cargo', 'Cód Nível Hierarq', 'Nível\nHierárquico',
                         'UO SUPERIOR', 'MATRÍCULA LÍDER', 'LÍDER', 'Centro de custo - aj CC', 'CARGO LÍDER', 'Sexo',
                         'Data de Nascimento', 'Idade', 'Grau de Instrução', 'PCD', 'Situação', 'Data de Rescisão',
                         'Motivo Desligamento', 'Tipo desligamento', 'Motivo Afastamento', 'Matrícula.1', 'CPF', 'RG',
                         'Grades Hay', 'Cód Praça', 'Praça', 'Região.1', 'Estrutura', 'ADMITIDOS', 'Bate recisao',
                         'Bate recisão mês','Ativo', 'área/loja - AJ','Status Generico','Bate Recindido','Centro de custo ajustado','Região',
                         'Entrevista','Tipo','Status final geral','área/loja - AJ ajustado','sigla - AJ ajustado','Diretoria - Ajustado',
                         'unidade negócio - aj ajustado','Novo cluster','setor']

# Selecionar as colunas desejadas
bd = bd[colunas_selecionadas]

# Renomear colunas
bd = bd.rename(columns={
    'Status final geral': 'Status Estabelecimento',
    'Diretoria final': 'Diretoria',
    'Unidade Resumida 24 com 23 (SEEEROO)': 'Unidade Resumida',
    'sigla - AJ ajustado': 'Sigla',
    'Centro de custo ajustado'  : 'Centro de Custo',
    'área/loja - AJ ajustado': 'área/loja',
    'Bate Recindido':'RECINDIDOS'
})
    
# Criar a nova coluna com base na condição
bd['COD_EMPRESA'] = bd['Cód Estabelecimento'].apply(lambda x: 1 if x < 500 else 2)

cargos_gestao = [
    'GTE PLANEJAMENTO ORCAMENTO OBRAS', 'COORD PROJETOS LOGISTICOS', 'COORD CONTROLES INTERNOS',
    'GTE LOJA II', 'SUPV VENDAS I', 'SUPV ACESSORIOS', 'SUPV RETAGUARDA', 'COORD PRODUCAO',
    'GTE LOJA I', 'GTE VM', 'COORD VISUAL MERCHANDISING', 'LIDER TELEVENDAS', 'COORD B2B', 'GTE B2B',
    'COORD TELEVENDAS', 'GERENTE BI', 'SUPV ADMINISTRATIVO', 'SUPV ESTOQUE',
    'SUPERVISOR OPERACAO LOGISTICA', 'COORD LOGISTICA', 'SUPERVISOR EMPILHADEIRA', 'COORD RECEBIMENTO FATURAMENTO',
    'GTE LOGISTICA', 'PRODUCT OWNER', 'PRODUCT MANAGER', 'DIRETOR DE TECNOLOGIA', 'SUPV LOJA COMPACTA',
    'GTE RELACIONAMENTO', 'COORD DADOS', 'COORD MARKETING CRM', 'GTE DESIGN TENDENCIAS SR',
    'COORD CADASTRO CONTEUDO', 'GTE DESIGN E TENDENCIAS', 'COORD SEO', 'GTE MARKETING PERFORMANCE',
    'SUPERVISOR QUALIDADE', 'GTE GERAL QUALIDADE TOTAL', 'GTE GERAL MARKETING', 'COORD MARKETING',
    'GTE BRANDING E COMUNICACAO', 'COORD BRANDING E COMUNICACAO', 'GTE GERAL CONTROLADORIA', 'COORD CONTABIL',
    'GTE CONTABILIDADE', 'COORD CUSTOS', 'ARQUITETO', 'GTE PROJETOS E OBRAS', 'GTE DESIGN ACESSORIOS',
    'GTE DESIGN MOVEIS', 'COORD GENTE E GESTÃO', 'GTE SR GESTAO DE PESSOAS', 'GTE PRODUTO II',
    'GTE GRUPO PRODUTO ACESSORIOS SR', 'GTE PRODUTO I', 'GTE GRUPO PRODUTO MOVEIS SR', 'GTE PLANEJAMENTO FINANCEIRO',
    'GTE TREINAMENTO LOJA', 'COORD TREINAMENTO E CONTEÚDO', 'GTE TESOURARIA', 'COORD CONTAS A PAGAR',
    'COORD DE CRÉDITO E COBRANÇA', 'COORD BANCO DADOS', 'SUPV INFRAESTRUTURA', 'COORD INFRAESTRUTURA TI',
    'SUPV OPERACAO TI', 'COORD SUPORTE APLICACOES', 'GTE INFRAESTRUTURA TI', 'SUPV NOC', 'SUPV QUALIDADE TI',
    'COORD DESENVOLVIMENTO & PROJETOS', 'GTE LOGISTICA REVERSA', 'COORD FISCAL', 'GTE FISCAL TRIBUTARIO',
    'DIRETOR COMERCIAL', 'GTE REGIONAL I', 'GTE GERAL VENDAS', 'SUPV VENDAS II',
    'PROCUREMENT CATEGORY MANAGER', 'COORD SOURCING', 'GTE GERAL PROCUREMENT', 'CHIEF SALES OFFICER',
    'CHIEF TECHNOLOGY OFFICER', 'CHIEF OPERATING OFFICER', 'CHIEF FINAN OFFICER', 'DIRETOR EXECUTIVO EXPANSÂO',
    'DIRETOR PRODUTO E ESTILO', 'GTE PLANEJAMENTO OPERACOES', 'GTE GERAL PLANEJAMENTO INTEGRADO',
    'GTE INDICADORES E ORCAMENTO', 'GTE PRICING', 'GTE IMPLEMENTACAO LOJAS', 'SUPERVISOR ENTREGAS',
    'COORD ENTREGAS', 'SUPV TRAFEGO', 'GTE ENTREGAS', 'SUPV SEGURANCA PATRIMONIO', 'COORD JURIDICO',
    'GTE GERAL JURIDICO', 'GTE INTELIGÊNCIA COMPRAS & GESTÃO FORNECEDORES', 'MÉDICO TRABALHO',
    'COORD SEGURANCA TRABALHO', 'GTE UX', 'GTE VISUAL MERCHANDISING SR', 'COORD ECOMMERCE', 'GTE SR ECOMMERCE',
    'COORD PROJETOS CX', 'SUPV RELACIONAMENTO', 'COORD RELACIONAMENTO', 'COORD EXPANSAO', 'SUPV MANUTENCAO',
    'GTE MULTICANAIS', 'SUPV OMNICHANNEL', 'COORD PLANEJAMENTO LOGISTICO', 'DIRETOR DE LOGÍSTICA',
    'SHIPPING PLANNING MANAGER', 'GTE PLANEJAMENTO ORCAMENTO OBRAS',
    'SUPERVISOR OPERACAO LOGISTICA II', 'SUPV LOJA', 'GTE GERAL LOGISTICA', 'GERENTE DADOS CRM',
    'COORD IDENTIDADE CRIATIVA', 'COORD QUALIDADE E REVERSA', 'COORD MARKETING PERFORMANCE',
    'DESIGNER PRODUTOS PL', 'DESIGNER PRODUTOS SR', 'COORD PLANEJAMENTO FINANCEIRO', 'GTE GERAL OPERACOES',
    'COORD REVERSA', 'GTE QUALIDADE E ENTREGAS', 'COORD MANUTENCAO FACILITIES', 'COORD SEGURANCA PATRIMONIO',
    'COORD QUALIDADE TOTAL', 'GTE PLANEJAMENTO LOGISTICO', 'DIRETOR DE LOGÍSTICA', 'COORD ORCAMENTO E COMPRAS'
]

niveis_hierarquicos_gestao = ['GERENCIA', 'COORDENACAO', 'SUPERVISAO', 'DIRETORIA', 'GERENCIA SR', 'GERENCIA GERAL',
                              'DIRETORIA EXECUTIVA', 'VP']

# Adicionar a nova coluna com base nas condições
bd['Gestão'] = bd.apply(lambda row: 'Gestão' if row['Cargo'] in cargos_gestao and row['Nível\nHierárquico'] in niveis_hierarquicos_gestao else 'Não Gestão', axis=1)

# Exibir os dados atualizados
print("Dados com a nova coluna 'Gestão':")
print(bd[['Cargo', 'Nível\nHierárquico', 'Gestão']])
############################ Salvar o arquivo

# Adicionar a nova coluna 'Mês aj' com base nas condições
bd['Mês aj'] = bd['Mês'].apply(lambda x: 'Janeiro' if x == 1 else ('Fevereiro' if x == 2 else ('Março' if x == 3 else
                           ('Abril' if x == 4 else ('Maio' if x == 5 else ('Junho' if x == 6 else
                           ('Julho' if x == 7 else ('Agosto' if x == 8 else ('Setembro' if x == 9 else
                           ('Outubro' if x == 10 else ('Novembro' if x == 11 else 'Dezembro')))))))))))

# Exibir os dados atualizados
print("Dados com a nova coluna 'Mês aj':")          
print(bd[['Mês', 'Mês aj']])



# geração

bd['AnoNascimento'] = bd['Data de Nascimento'].dt.year

def definir_geracao(ano):
    if 1928 <= ano <= 1945:
        return 'Silent Generation (1928 - 1945)'
    elif 1946 <= ano <= 1964:
        return 'Baby Boomer (1946 - 1964)'
    elif 1965 <= ano <= 1980:
        return 'Geração X (1965 - 1980)'
    elif 1981 <= ano <= 1996:
        return 'Millennials (1981 - 1996)'
    elif 1997 <= ano <= 2012:
        return 'Geração Z (1997 - 2012)'
    else:
        return 'Geralção Alfa'

bd['Geracao'] = bd['AnoNascimento'].apply(definir_geracao)


print (bd['Geracao'])


bd['N_Aprendiz'] = bd['Cargo'].str.upper().str.contains('APRENDIZ')



associacoes = {
    'Alessandra Nasser': ['BCS', 'BSB', 'BSI', 'CPB', 'CTB', 'FLO', 'GNA', 'JOI', 'LON', 'POA', 'SIP'],
    'Carol Manfredi': ['ALP', 'CGR', 'D&D', 'IBI', 'MKP', 'PIN', 'SCZ', 'SOC', 'STO'],
    'Mari Seghese': ['ABC', 'CAM', 'CSP', 'HIG', 'MOC', 'MTE', 'POM', 'RPI', 'SJC', 'SJP'],
    'Rosana matos': ['AJU', 'BEL', 'JPA', 'MCO', 'MPN', 'NTL', 'RMF', 'RMS', 'SLZ', 'SSA', 'SSH', 'TRS'],
    'W Carla': ['BHO', 'BRT', 'BTF', 'CBA', 'COP', 'NIT', 'TJC', 'VTA']
}


def regional_correto(Sigla):
    for nome, siglas_associadas in associacoes.items():
        if Sigla in siglas_associadas:
            return nome
    return None



# Adicionando uma coluna ao DataFrame com os nomes correspondentes às siglas
bd['Regional_correto'] = bd['Sigla'].apply(regional_correto)

# Exibindo o DataFrame resultante
print(bd)

bd['N_PCD'] = bd['PCD'].str.upper().str.contains('S')


# Configurar a janela Tkinter para salvar o arquivo
root_save = Tk()

# Abrir a caixa de diálogo para salvar o arquivo Excel
caminho_salvar = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])

# Fechar a janela Tkinter após a seleção do local para salvar o arquivo
root_save.destroy()

# Salvar o DataFrame como um novo arquivo Excel
bd.to_excel(caminho_salvar, index=False)

print(f"O arquivo foi salvo em: {caminho_salvar}")


print('Em caso de erro contatar: ###')
