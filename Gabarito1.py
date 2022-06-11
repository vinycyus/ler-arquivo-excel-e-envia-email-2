import pandas as pd
import smtplib

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from email.mime.base import MIMEBase
from email import encoders

# importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

# faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)
# ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

##corpo do email
relatorio_html= f'''
<p>Prezados,</p>

<p>Segue o Relatório de Vendas por cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer dúvida estou à disposição.</p>

<p>Att.,</p>
<p>vinicius</p>
'''
# enviar um email com o relatório

#S M T P - Simple Mail transfer protocol 
#Para criar o servidor e enviar o e-mail

#1- STARTAR O SERVIDOR SMTP --variaveis
host = "smtp.gmail.com"
port = "587"
login = "viniciussowza16@gmail.com" #login do o email remetente
senha = "bicicleta28" #senha do email remetente
destinario="matheuscnn2@gmail.com"
#Dando start no servidor 
server = smtplib.SMTP(host,port)
server.ehlo()
server.starttls()
server.login(login,senha)

#montando e-mail 
email_msg = MIMEMultipart()
email_msg['From'] = login #rementente
email_msg['To'] = destinario  #destinatario
email_msg['Subject'] = "<b>Meu e-mail enviado por pitão</b>" #assunto do email
email_msg.attach(MIMEText(relatorio_html,'html'))


#Abrimos o arquivo em modo leitura e binary 
#cam_arquivo = "C:\\pasta4\\Email\\teste.py" 
#attchment = open(cam_arquivo,'rb')

#Lemos o arquivo no modo binario e jogamos codificado em base 64 (que é o que o e-mail precisa )
#att = MIMEBase('application', 'octet-stream')
#att.set_payload(attchment.read())
#encoders.encode_base64(att)

#ADCIONAMOS o cabeçalho no tipo anexo de email 
#att.add_header('Content-Disposition', f'attachment; filename=teste.py')
#fechamos o arquivo 
#attchment.close()

#colocamos o anexo no corpo do e-mail 
#email_msg.attach(att)

#3- ENVIAR o EMAIL tipo MIME no SERVIDOR SMTP 
server.sendmail(email_msg['From'],email_msg['To'],email_msg.as_string())
server.quit()