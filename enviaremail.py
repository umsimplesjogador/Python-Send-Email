import win32com.client as win32

#criar a integração com o outlook
outlook = win32.Dispatch('outlook.application')

#criar um email
email = outlook.CreateItem(0)

#variáveis que podem ser modificadas de acordo com a necessidade
faturamento = 1500
qtde_produtos = 10
ticket_medio = faturamento / qtde_produtos

#configurar as informações do seu email
email.To = 'destino; destino2'
email.Subject = 'assunto'
email.HTMLBody = f'''
<p>Olá pessoa, aqui é o texto que você deseja escrever no corpo do email!</p>

<p>O faturamento da empresa foi de R$ {faturamento}</p>
<p>Vendemos {qtde_produtos} produtos</p>
<p>O Ticket Médio foi de R$ 

<p>Finalizando!</p>
<p>Fábio Silva</p>
'''

#colocar o local do arquivo para anexar ao email
'''anexo = 'C://Users/ap44/Downloads/arquivo.xlsx'
email.Attatchments.Add(anexo)'''

email.Send()
print('Email Enviado')
