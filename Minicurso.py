import pandas as pd
import win32com.client as win32

# importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)
print('-' * 50)
# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)
# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)
# ticket médio por produto em cada loja
ticketmedio = (faturamento['Valor Final']/quantidade['Quantidade']).to_frame()
ticketmedio = ticketmedio.rename(columns={0: 'Ticket Médio'})
print(ticketmedio)
print('-' * 50)
# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
# colocar endereço de e-mail que deseja enviar
mail.To = 'xxxxxxx'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
<p>Olá,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticketmedio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Atenciosamente,</p>
<p>Bárbara</p>'''

mail.Send()