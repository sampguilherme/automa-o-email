import pandas as pd
import smtplib as smtp
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import time
import pwinput
import base64

# Ler planilha excel (Trata como banco de dados)
excelSheet = pd.read_excel("testesFinais.xlsx")

# Ler as informações de login do usuário
user = input("Digite seu nome: ")
login = input("Digite o e-mail para login: ")
password = pwinput.pwinput("Digite a senha: ")

# Realiza conexão do email com servidor SMTP
server = smtp.SMTP('smtp-mail.outlook.com', 587)
server.ehlo()
server.starttls() # Torna a conexão segura
server.login(login, password)


with open('ASS_GUILHERME_SAMPAIO.png', "rb") as image_file:
    encoded_image = base64.b64encode(image_file.read()).decode('utf-8')

# Estrutura do e-mail 

body = f"""
    <html>
        <body>
            <p>Boa tarde,</p>
            <p>Meu nome é {user} trabalho na prospecção de novos canais para a Brico Bread.</p>
            <p>Com mais de 80 anos de história na produção de pães, a Brico Bread é a maior e melhor fabricante de pré-assados e ultracongelados da América do Sul. Nascemos da excelência de fazer pão com inovação, tecnologia de ponta. Além de garantir qualidade e excelência em seus produtos, a Brico oferece diversidade, confiança, escalabilidade e, principalmente, disponibilidade para parcerias e conexões.</p>
            <p>Trabalhamos com uma linha de pré-assados e ultracongelados que dispensam qualquer tipo de equipamento técnico e mão de obra especializada para finalização.</p>
            <p>Nós encontramos a sua empresa, e vimos um grande potencial para uma parceria. Por esse motivo, gostaríamos de marcar uma reunião com vocês e com a nossa diretoria.</p>
            <p>Me confirme qual a data e horário teriam disponibilidade, assim consigo enviar o invite da call.</p>
            <p>Em tempo: <a href="https://publicbrico.s3.sa-east-1.amazonaws.com/brico.pdf"> Clique aqui</a> para acessar nosso catálogo. Temos uma variedade de até 50 tipos de produtos.</p>
            <p>Fico no aguardo de um breve retorno.</p>
            <p>Obrigada.</p>
            <img src="data:image/png;base64,{encoded_image}" alt="Assinatura">
        </body>
    </html>
"""


# Loop para envio de e-mails
emails_send = 0
for i in range(excelSheet.__len__()):
    if excelSheet.__len__() <= 0: # Condição para o código parar se a planilha do excel for zerada
        break
    if emails_send >= 100: # Condição para o script mandar uma determinada quantidade de emails e pausar por um determinado tempo
        emails_send = 0 
        excelSheet.to_excel("teste.xlsx", index=False)
        time.sleep(300) 
    email_msg = MIMEMultipart()
    name = ""
    if excelSheet.loc[i, "Nome Fantasia"] == str:
        name = excelSheet.loc[i, "Nome Fantasia"]
    else:
        name = excelSheet.loc[i, "Razao Social"]
    email_msg['From'] = login
    email_msg['To'] = excelSheet.loc[i, "E-mail"]
    email_msg['Subject'] = "Parceria Brico Bread c/ " + name
    email_msg.attach(MIMEText(body, 'html'))
    print(excelSheet.loc[i, "E-mail"], name)
    server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
    emails_send = emails_send + 1
    excelSheet = excelSheet.drop(i) # Deleta a linha do e-mail enviado
    
    
excelSheet.to_excel("teste.xlsx", index=False) # Atualiza na planilha todas as linhas que foram deletadas 
server.quit() # Desconecta do servidor SMTP
