### Fazendo uma análise de vendas das lojas dos principais Shoppings, enviar um email automático e adicionar os dados no database ###

"""
Informações fornecidas:
Código Venda
Data
ID Loja
Produto
Quantidade
Valor Unitário
Valor Final
"""

""" Calcular """
# Faturamento por loja
# Quantidade de produtos vendidos por loja
# Ticket médio por produto em cada loja (Faturamento/Quantidade)

# Enviar um e-mail com o relatório

### PASSOS ###
# 1o - Importar Base de dados com as vendas
# 2o - Visualizar a base
# 3o - Aplicar os tratamentos dito acima 
# 4o - Enviar Email automatico
# 5o - Adicionar dados na base de dados

##############################
### Instalações Essenciais ###
##############################

# pip install pandas
# pip install openpyxl
# pip install pywin32
# pip install mysqlconnector

##############################
### Bibliotecas Essenciais ###
##############################

import pandas as pd
import openpyxl
import win32com.client as win32
import mysql.connector
from mysql.connector import Error
import config  # Importa as configurações do arquivo config.py

##############################
###### Funções Projeto #######
##############################

# Ler arquivo excel
tabela_vendas = pd.read_excel('Vendas_lojas_shopping.xlsx')

# Visualizar base de dados completa
pd.set_option('display.max_columns', None)
print("Antes do Faturamento: ")
print("-" * 30)
print(tabela_vendas)

""" Calcular Faturamento """
# Agrupar todas as lojas para ficar 1 só vez cada loja no dataframe e somar as colunas de valor final, .sum para somar
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print("-" * 30)  
print("Depois do Faturamento: ")
print(faturamento)

""" Quantidade de Produtos vendidos por loja """
quantidade_vendas = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()  # .sum para Somar as quantidades de vendas por loja
print("-" * 30)
print("Depois da Quantidade: ")
print(quantidade_vendas)

""" Calcular Ticket Médio """
# Dividir o Faturamento (Valor Final) pela Quantidade
ticket_medio = faturamento['Valor Final'] / quantidade_vendas['Quantidade']
ticket_medio = ticket_medio.to_frame(name='Ticket Médio')  # Convertendo para DataFrame
print("-" * 30)
print("ticket_medio:")
print(ticket_medio)

# Juntar todas as colunas em um único DataFrame
df_vendas_final = faturamento.join(quantidade_vendas).join(ticket_medio)
print("-" * 30)
print("DataFrame Final:")
print(df_vendas_final)

# Alterando o nome das colunas para caber no mysql
df_vendas_final.rename(columns={
    'ID Loja': 'ID_Loja',
    'Valor Final': 'Valor_Final',
    'Ticket Médio': 'Ticket_Médio'
}, inplace=True)

""" Adicionar dados na base de dados """

# OBS: AQUI VOCÊ PRECISARÁ CRIAR UM DATABASE NO MYSQL E COLOCAR AS CREDENCIAIS NO ARQUIVO .env (removido pelo .gitignore)

try:
    # Conectar ao banco de dados
    conexao_db = mysql.connector.connect(
        host=config.DB_HOST,
        database=config.DB_DATABASE,
        user=config.DB_USER,
        password=config.DB_PASSWORD
    )

    if conexao_db.is_connected():
        print("Conexão ao banco de dados estabelecida com sucesso!")
        cursor = conexao_db.cursor()

        # Substituindo valores nulos por 0 ou "" para não ficar com valores NaN
        df_vendas_final.fillna({
            'ID_Loja': "",
            'Valor_Final': 0,
            'Quantidade': 0,
            'Ticket_Médio': 0,
        }, inplace=True)

        # Inserindo valores nas querys
        query_insert_vendas = """
        INSERT INTO TB_VENDAS_LOJAS
        (ID_Loja, Valor_Final, Quantidade, Ticket_Médio)
        VALUES (%s, %s, %s, %s)
        ON DUPLICATE KEY UPDATE
        Valor_Final = VALUES(Valor_Final), 
        Quantidade = VALUES(Quantidade), 
        Ticket_Médio = VALUES(Ticket_Médio)
        """

        for row in df_vendas_final.itertuples():
            cursor.execute(query_insert_vendas, (row.Index, row.Valor_Final, row.Quantidade, row.Ticket_Médio))

        conexao_db.commit()
        print("Dados inseridos com sucesso no banco de dados!")

    cursor.close()
    conexao_db.close()

except Error as e:
    print(f"Erro ao conectar ao banco de dados: {e}")

# Enviar e-mail 
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'testeautomacaopython1@outlook.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade_vendas.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}
.

'''

mail.Send()

print('Email Enviado')

