import csv, smtplib, os, openpyxl
import pandas as pd
from colorama import Fore, Back, Style
from email.mime.text import MIMEText

# Método para enviar email
def send_email(subject, body, sender, recipients, password):
    body_text = '\n'.join(body)
    msg = MIMEText(body_text)
    msg['Subject'] = subject
    msg['From'] = sender
    msg['To'] = ', '.join(recipients)  

    # Configurando o servidor SMTP
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp_server:
        smtp_server.login(sender, password)
        smtp_server.sendmail(sender, recipients, msg.as_string())
    print("Mensagem enviada!")

relatorio_csv = "relatorio_esteiras_com_problemas.csv"

with open(relatorio_csv, 'w', newline='', encoding='utf-8') as arquivo_relatorio_csv:
                spamwriter = csv.writer(arquivo_relatorio_csv)
                spamwriter.writerow(['Date', 'Time', 'esteira0', 'esteira1', 'esteira2'])
                
# Ler o arquvio .csv
with open('ESP8266_Receiver - Sheet1.csv', newline='') as arquivo_csv:
    conteudo_arquivo_csv = csv.reader(arquivo_csv, delimiter=',')
    next(conteudo_arquivo_csv)  # Ignorar primeira linha do arquivo_csv
    body = []
    
    for linha_arquivo_csv in conteudo_arquivo_csv:  # Percorre cada linha da lista

        # Verifica os níveis críticos de cada esteira e adiciona ao corpo do e-mail
        if int(linha_arquivo_csv[2]) < 5:  # esteira 0
            body.append(f"A esteira 0 no dia {linha_arquivo_csv[0]} no tempo: {linha_arquivo_csv[1]}\nestá abaixo do nível crítico!!\n")

            with open(relatorio_csv, 'a', newline='', encoding='utf-8') as arquivo_relatorio_csv:
                spamwriter = csv.writer(arquivo_relatorio_csv)
                spamwriter.writerow([linha_arquivo_csv[0], linha_arquivo_csv[1], linha_arquivo_csv[2], '-', '-'])

        if int(linha_arquivo_csv[3]) < 250:  # esteira 1
            body.append(f"A esteira 1 no dia {linha_arquivo_csv[0]} no tempo: {linha_arquivo_csv[1]}\nestá abaixo do nível crítico!!\n")

            with open(relatorio_csv, 'a', newline='', encoding='utf-8') as arquivo_relatorio_csv:
                spamwriter = csv.writer(arquivo_relatorio_csv)
                spamwriter.writerow([linha_arquivo_csv[0], linha_arquivo_csv[1], '-', linha_arquivo_csv[3], '-'])

        if int(linha_arquivo_csv[4]) < 25000:  # esteira 2
            body.append(f"A esteira 2 no dia {linha_arquivo_csv[0]} no tempo: {linha_arquivo_csv[1]}\nestá abaixo do nível crítico!!\n")

            with open(relatorio_csv, 'a', newline='', encoding='utf-8') as arquivo_relatorio_csv:
                spamwriter = csv.writer(arquivo_relatorio_csv)
                spamwriter.writerow([linha_arquivo_csv[0], linha_arquivo_csv[1], '-', '-', linha_arquivo_csv[4]])

subject = "Detalhamento da Esteira"
sender = "" # Utilize o Email da sua preferência
password = ""  # Use sua senha de app gerada pelo Google
recipients = ["", ""] # Coloque o email para quem será enviado

send_email(subject, body, sender, recipients, password)

# Transformar o arquivo "relatorio_esteiras_com_problemas.csv" em um .xlsx
novo_conjunto_dados = pd.read_csv(relatorio_csv)

# Salva como arquivo Excel com extensão correta
relatorio_excel = "relatorio_esteiras_com_problemas.xlsx"
novo_conjunto_dados.to_excel(relatorio_excel, index=False)

print(f"Relatório salvo como: {relatorio_excel}")