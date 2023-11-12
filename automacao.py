import pandas as pd

#Lê a planilha do excel
df = pd.read_excel("teste.xlsx")





for i in range(df.__len__()):
    print(df.loc[i])
    df = df.drop(i) 
    
df.to_excel("teste.xlsx", index=False)

import smtplib as smtp
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText


login = "example@hotmail.com"
password = "example1213"

# Realiza a conexão do email

server = smtp.SMTP('smtp-mail.outlook.com', 587)

server.ehlo()
server.starttls()
server.login(login, password)


body = "Olá, este é um e-mail de teste!"

email_msg = MIMEMultipart()
email_msg['From'] = login
email_msg['To'] = "example@hotmail.com"
email_msg['Subject'] = "Testando"
email_msg.attach(MIMEText(body, 'Plain'))

server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string()) 

server.quit()