import pandas as pd
import win32com.client as win32

# Sempre pensar na lógica em português e depois passar ela para linguagem Python

# PASSO A PASSO:
# Importar base de dados
tabela_vendas = pd.read_excel("Vendas.xlsx")

# Visualizar base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)


# Faturamento por loja

# Filtrar colunas para serem visualizadas
# Agrupar as lojas e somar faturamento
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby(
    'ID Loja').sum()
print(faturamento)

# Quantidade de produtos vendidos por loja
qtd_produtos_vendidos = tabela_vendas[[
    'ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(qtd_produtos_vendidos)

print('-' * 50)
# Ticket médio por produto em cada loja => faturamento / qtd vendida do produto
ticket_medio = (faturamento['Valor Final'] /
                qtd_produtos_vendidos['Quantidade']).to_frame()
# quando faz uma operação matemática entre uma coluna e outra, o resultado não é retornado em tabela e sim em um amontoado de dados
# para transformar em tabela, usa-se o to_frame

print(ticket_medio)

# Enviar um email com o relatório

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'pythonimpressionador@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{qtd_produtos_vendidos.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>Vanessa</p>
'''

mail.Send()

print('Email Enviado')
