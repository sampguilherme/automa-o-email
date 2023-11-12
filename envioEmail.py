import pandas as pd
import smtplib as smtp
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Ler planilha excel (Trata como banco de dados)
excelSheet = pd.read_excel("teste.xlsx")

# Ler as informações de login do usuário
login = input("Digite o e-mail para login: ")
password = input("Digite a senha: ")

# Realiza conexão do email com servidor SMTP
server = smtp.SMTP('smtp-mail.outlook.com', 587)

server.ehlo()
server.starttls()
server.login(login, password)

# Estrutura do e-mail 
body = "Lorem ipsum dolor sit amet. Ut natus perspiciatis nam culpa quidem ut maxime dicta eum fugit alias non dolorem excepturi. Non commodi facilis cum enim distinctio est quis explicabo non vero voluptatum. 33 veritatis voluptatem sit enim dignissimos aut soluta laudantium et quia dolorum et Quis voluptatem. Ex sunt excepturi nam quibusdam cumque aut itaque libero ab adipisci Quis ut dolore voluptatem nam quaerat reprehenderit."

email_msg = MIMEMultipart()
email_msg['From'] = login
email_msg.attach(MIMEText(body, 'Plain'))
# Loop para envio de e-mails
for i in range(excelSheet.__len__()):
    if i >= 2:
        break
    emailTo = excelSheet.loc[i, "E-mail"]
    name = excelSheet.loc[i, "Nome"]
    print(type(emailTo), type(name))
    emailTo = excelSheet.loc[i, "E-mail"]
    name = excelSheet.loc[i, "Nome"]
    email_msg['To'] = str(emailTo)
    email_msg['Subject'] = "Olá, " + name
    server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())
    excelSheet = excelSheet.drop(i) # Deleta a linha do e-mail enviado
    

excelSheet.to_excel("teste1.xlsx", index=False) # Atualiza na planilha todas as linhas que foram deletadas
server.quit() # Desconecta do servidor SMTP 