# -*- coding: utf-8 -*-
import pandas as pd
import smtplib
import email.message


tabela_vendas = pd.read_excel('Vendas.xlsx')
pd.set_option('display.max_rows', None)

#filtra as tabelas da loja e do valaor final, e para evitar a repetição de cada shoping, utilizo o groupby.sum()
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()

print(faturamento)

print('-'*50)

quantidade_produto = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()

print(quantidade_produto)

print('-'*50)

ticket_medio = (faturamento['Valor Final']/quantidade_produto['Quantidade']).to_frame()

ticket_medio = ticket_medio.rename(columns={0:'Ticket Médio'})

print(ticket_medio)



def enviar_email():
    corpo_email = f"""
    
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Document</title>
    </head>
    <body>
        <p> Prezados, segue o relatório de vendas. </p>
    
        <p> Faturamento:</p>
        <p> {faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}</p>
    
        <p> Quantidade de produtos vendidos:</p>
        <p> {quantidade_produto.to_html()}</p>
    
        <p> Ticket médio dos produtos em cada loja:</p>
        <p> {ticket_medio.to_html(formatters={'Ticket Médio':'R${:,.2f}'.format})}</p>
    </body>
    </html>
    
    
    
    
    """

    msg = email.message.Message()
    msg['Subject'] = "Relatório"
    msg['From'] = 'seuemail@gmail.com'
    msg['To'] = 'emaildestinatario@outlook.com'
    password = 'app pass do G-mail'
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')

enviar_email()