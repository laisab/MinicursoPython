import pandas as pd
import win32com.client

# Importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar a base de dados
# (opcao, valor)
pd.set_option('display.max_columns', None)
print(tabela_vendas)

print('-' * 50)
# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# Colunas da tabela fornecidas em forma de lista
# tabela_vendas[['ID Loja', 'Valor Final']]
# A coluna que serve para agrupar deve ser passada na lista

print('-' * 50)
# Quantidade de produtos vendidos por loja
qtde = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(qtde)

print('-' * 50)
# Ticket médio por produto em cada loja
# Não tem como fazer (faturamento / qtde) porque ambos são tabelas
ticket_medio = (faturamento['Valor Final']/qtde['Quantidade']).to_frame()
# to_frame(): transforma em tabela
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

# Enviar um email com o relatório
# Conexão com o Outlook instalado
outlook = win32com.client.Dispatch('outlook.application')
# Criação do email
outlook_item = 0x0
mail = outlook.CreateItem(outlook_item)
# Configuração do email
mail.Subject = 'Relatório de Vendas por Loja'
mail.To = 'laisborges@alunos.utfpr.edu.br'
mail.HTMLBody = f'''
<p>Prezados, segue o Relatório de Vendas por cada loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R$ {:,.2f}'.format})}

<p>Quantidade vendida:</p>
{qtde.to_html()}

<p>Ticket médio dos produtos em cada loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R$ {:,.2f}'.format})}

<p>Qualquer dúvida, estou à disposição.</p>

<p>Atte., Lais</p>'''

mail.Send()
