# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.

Creator 78154 Wendel Ventura 
"""

import pandas as pd
from tkinter import Tk, filedialog

def dt_bd_promo():
    # Configurar a janela Tkinter
    root2 = Tk()

    try:
        # Abrir a caixa de diálogo para seleção do arquivo
        caminho_bd_promo = filedialog.askopenfilename(title="Selecione o arquivo Excel")

        # Fechar a janela Tkinter após a seleção do arquivo
        root2.destroy()
        
        # Ler o arquivo CSV excluindo a primeira linha
        bd_promo = pd.read_csv(caminho_bd_promo, sep=';', encoding='ISO-8859-1',skiprows=1)

        # Converter a coluna 'Data da Operacao' para datetime, especificando o formato
        bd_promo["Data da Movimentação"] = pd.to_datetime(bd_promo["Data da Movimentação"], format='%d/%m/%Y', errors='coerce')

        # Adicionar coluna de ano
        bd_promo['Ano'] = bd_promo['Data da Movimentação'].dt.year

        # Adicionar coluna de mês
        bd_promo['Mês'] = bd_promo['Data da Movimentação'].dt.month
        
        # Adicionar chave_CC
        bd_promo['Chave_CC'] = bd_promo['Código da Empresa de Destino'].astype(str) + bd_promo['Código do Centro de Custo de Destino'].astype(str) + bd_promo['Código do Estabelecimento de Destino'].astype(str)

        # Adicionar chave_data
        bd_promo['Chave_data'] = bd_promo['Ano'].astype(str) + bd_promo['Mês'].astype(str) + bd_promo['Código do Centro de Custo de Destino'].astype(str)

        bd_promo['Matricula'] = bd_promo.apply(lambda row: row['Matrícula de Origem'] if row['Matrícula de Destino'] == '' else row['Matrícula de Origem'], axis=1)

    
        # Exibir os dados
        print("Lendo os dados do arquivo:")
        print(bd_promo)

        return bd_promo

    except Exception as e:
        print(f"Ocorreu um erro: {e}")
        return None

# Chamar a função
bd_promo = dt_bd_promo()

# Configurar a janela Tkinter para salvar o arquivo
root_save = Tk()

# Abrir a caixa de diálogo para salvar o arquivo Excel
caminho_salvar = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])

# Fechar a janela Tkinter após a seleção do local para salvar o arquivo
root_save.destroy()

# Salvar o DataFrame como um novo arquivo Excel
bd_promo.to_excel(caminho_salvar, index=False)

print(f"O arquivo foi salvo em: {caminho_salvar}")


print('Em caso de erro contatar: wendel.almeida@tokstok.com.br')



