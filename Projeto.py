import pandas as pd
import win32com.client as win32

tabela_vendas = pd.read_excel('Vendas.xlsx')

pd.set_option('display.max_columns', None)  # Nesse caso não é necessario

# Faturamento da Loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby(
    'ID Loja').sum()

# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()

print('-' * 50)

# ticket medio da loja
ticket_medio = (faturamento['Valor Final'] /
                quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
# coluna ficava como 0, tambem pode ser trocada no to_frame acima
print(ticket_medio)


# Enviar email, precisa do outlook instalado localmente (versão paga)
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

# (formatters={'Valor Final': 'R${:,.2f}'.format})} a virgula (casas dos milhares), o ponto (centavos, 2f é 2 casas)

mail.To = 'teste@hotmail.com'
mail.Subject = 'Teste Relatorio'
mail.HTMLBody = f'''
<p> Olá amigo, teste relatorio </p>

<p> Faturamento: </p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p> Quantidade: </p>
{quantidade.to_html()}

<p> Ticket medio: </p>
{ticket_medio.to_html()}

'''
mail.send()
