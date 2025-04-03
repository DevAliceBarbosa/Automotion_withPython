#Nesse projeto de automação aprenderei como a partir de uma planilha do excel fazer o

# 0- Visualizar a base de dados

# 1- Faturamento por loja

# 2- Quantidade de produtos vendidos por loja

# 3- Ticket médio por produto em cada loja

# 4- Enviar um email com o relatório

#---------------------//----------------------------
#Iniciando o projeto de automação para enviar um relatório por email
#---------------------//----------------------------

#Importações necessarias para o nosso projeto
import pandas as pd
import win32com.client as win32


#Importar a base de dados (Planilha do excel)
tabelas_vendas = pd.read_excel('Vendas.xlsx')

# Fazendo o item 0 (Visualizar a base de dados)
pd.set_option('display.max_columns', None) #Serve para não colocar limite na visualização de colunas quando der print na planilha
print(tabelas_vendas)

#Metodo Filtrar colunas de uma tabela com panda
#tabelas_vendas[['ID Loja', 'Valor Final']]
#Metodo tabela.groupby (Serve para agrupar)
#tabelas_vendas.groupby('ID Loja').sum() #vai agrupar as lojas e somar as outras colunas

# Fazendo o item 1 (Faturamento por loja) utilizando os metodos de cima
faturamento = tabelas_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)


#Fazendo o item 2 (Quantidade de produtos vendidos por loja)
quantidadeProdutos = tabelas_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidadeProdutos)
print('-' * 50)


#Fazendo o item 3 (Ticket medio)
ticket_medio = (faturamento['Valor Final'] / quantidadeProdutos['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)


#Fazendo o item 4 (Enviando o email com o relatorio)


outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To =  'xxx@xxx.com'
mail.Subject = 'XXXX'
mail.HTMLBody = f'''
<h2>Test </h2>
<p>Esse email aqui não tem nenhum uso, é apenas para eu saber se o minicurso de automação de processos com Python deu certo...</p>

<p>Utilizei uma planilha do excel e realizei a manipulação dos dados desse relatório com um código em Python
Isso é bastante util para grandes empresas que precisam agilizar o serviço.</p>

<p><b>Faturamento</b></P
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p><b>Quantidade Vendida</b></p>
{quantidadeProdutos.to_html()}

<p><b>Ticket Médio dos Produtos em cada Loja</b></p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Att...</p>
<p>Alice :)</p>
'''
mail.Send()
print("Email enviado")



