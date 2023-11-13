import pandas as pd
import smtplib as smtp
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
import time
import pwinput

# Ler planilha excel (Trata como banco de dados)
excelSheet = pd.read_excel("teste.xlsx")

# Ler as informações de login do usuário
login = input("Digite o e-mail para login: ")
password = pwinput.pwinput("Digite a senha: ")

# Realiza conexão do email com servidor SMTP


# Estrutura do e-mail 

body = """
    <html>
        <body>
            <p>Boa tarde,</br>
            Meu nome é Guilherme trabalho na prospecção de novos canais para a Brico Bread.</br>
            Com mais de 80 anos de história na produção de pães, a Brico Bread é a maior e melhor fabricante de pré-assados e ultracongelados da América do Sul. Nascemos da excelência de fazer pão com inovação, tecnologia de ponta. Além de garantir qualidade e excelência em seus produtos, a Brico oferece diversidade, confiança, escalabilidade e, principalmente, disponibilidade para parcerias e conexões.</br>
            Trabalhamos com uma linha de pré-assados e ultracongelados que dispensam qualquer tipo de equipamento técnico e mão de obra especializada para finalização.</br>
            Nós encontramos a sua empresa, e vimos um grande potencial para uma parceria. Por esse motivo, gostaríamos de marcar uma reunião com vocês e com a nossa diretoria.</br>
            Me confirme qual a data e horário teriam disponibilidade, assim consigo enviar o invite da call.</br>
            Em tempo, segue o nosso catálogo de produtos: https://publicbrico.s3.sa-east-1.amazonaws.com/brico.pdf temos uma variedade de até 50 tipos de produtos.</br>
            Fico no aguardo de um breve retorno.</br>
            Obrigada.</p>
            </br>
            <img src="cid:0" alt="Assinatura">
        </body>
    </html>
"""

email_msg = MIMEMultipart()
email_msg['From'] = login
email_msg.attach(MIMEText(body, 'html'))

# Adicionar assinatura no e-mail
signature = open('ASS_GUILHERME_SAMPAIO.png', "rb")
msg_signature = MIMEImage(signature.read())
signature.close()
msg_signature.add_header('Content-ID', '0')


# Loop para envio de e-mails
emails_send = 0
for i in range(excelSheet.__len__()):
    server = smtp.SMTP('smtp-mail.outlook.com', 587)
    server.ehlo()
    server.starttls() # Torna a conexão segura
    server.login(login, password)

    if excelSheet.__len__() <= 0: # Condição para o código parar se a planilha do excel for zerada
        break
    if emails_send >= 1: # Condição para o script mandar uma determinada quantidade de emails e pausar por um determinado tempo
        emails_send = 0 
        excelSheet.to_excel("teste.xlsx", index=False)
        time.sleep(10) 
    '''emailTo = ""
    emailTo = excelSheet.loc[i, "E-mail"]
    name = excelSheet.loc[i, "Nome"]'''
    email_msg['To'] = excelSheet.loc[i, "E-mail"]
    email_msg['Subject'] = "Parceria Brico Bread c/ " + excelSheet.loc[i, "Nome"]
    print(excelSheet.loc[i, "E-mail"], excelSheet.loc[i, "Nome"])
    server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
    emails_send = emails_send + 1
    excelSheet = excelSheet.drop(i) # Deleta a linha do e-mail enviado
    server.quit() # Desconecta do servidor SMTP
    
excelSheet.to_excel("teste.xlsx", index=False) # Atualiza na planilha todas as linhas que foram deletadas 

