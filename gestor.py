# bibliotecas e pacotes
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import smtplib
from email.mime.base import MIMEBase
from email import encoders
import os
import pandas as pd
from reportlab.pdfgen import canvas
import gspread
from google.oauth2.service_account import Credentials
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
import time
from datetime import datetime, timedelta

#Função para criar o PDF
def cria_PDF(i):
    # Criando o PDF
    pdf_filename = f"Autorização de Viagem Reforço_{i}.pdf"
    c = canvas.Canvas(pdf_filename)
    c.setFont("Helvetica", 12)
    imagem_path = "C:/Users/Administrador/Downloads/SUMOB.png"
    pos_x = 100
    pos_y = 750
    largura = 150
    altura = 60
    imagem = ImageReader(imagem_path)
    c.drawImage(imagem, pos_x -50, pos_y -20, largura, altura)
    c.drawString(100 - 50, 720 -20, "                                              Autorização de Viagem Reforço")
    c.drawString(100 -50, 700 -20, "")
    c.drawString(100 -50, 680 -20, "")
    c.drawString(100 -50, 660 -20, f"Data e Hora do Lançamento: {carimbo}")
    c.drawString(100 -50, 640 -20, f"E-mail: {email}")
    c.drawString(100 -50, 620 -20, f"Agente: {Agente}")
    c.drawString(100 -50, 600 -20, f"BT: {BT}")
    c.drawString(100 -50, 580 -20, f"Motivo: {motivo}")
    c.drawString(100 -50, 560 -20, f"Data da viagem: {data}")
    c.drawString(100 -50, 540 -20, f"Sistema: {sistema}")
    c.drawString(100 -50, 520 -20, f"Linha: {linha}")
    c.drawString(100 -50, 500 -20, f"Sublinha: {sublinha}")
    c.drawString(100 -50, 480 -20, f"PC: {partidasPC}")
    c.drawString(100 -50, 460 -20, f"Número de Ordem do Veículo: {NºdeOrdem}")
    c.drawString(100 -50, 440 -20, f"Placa: {Placa}")
    c.drawString(100 -50, 420 -20, f"Código da Viagem: {CodViagem}")
    c.drawString(
        100 -50, 400 -20, f"Horário de Abertura da Viagem no CITGIS {HorarioAberturaMCO}"
    )
    c.drawString(100, 300, "")
    c.drawString(100 -50, 290 -20, "Autorizado por:")
    c.drawString(100, 270, "")
    c.drawString(100 -50, 250 -20, "                                                  Gerência/BT")
    c.showPage()
    c.save()
    
    return pdf_filename


def enviar_email(subject, to, body, files):
    print(f"Enviando PDF {files}")
    msg = MIMEMultipart()
    msg["Subject"] = subject
    msg["From"] = "Inserir E-mail"
    msg["To"] = to
    password = "Inserir Senha"
    msg.attach(MIMEText(body, "html"))

    with open(files, "rb") as fil:
        part = MIMEApplication(fil.read(), Name=os.path.basename(files))
    part["Content-Disposition"] = 'attachment; filename="%s"' % os.path.basename(files)
    msg.attach(part)

    s = smtplib.SMTP("smtp.gmail.com:587")
    s.starttls()
    s.login(msg["From"], password)
    s.sendmail(msg["From"], [msg["To"]], msg.as_string().encode("utf-8"))
    print("Email enviado")


#Criar uma API no Google Console para inserir abaixo
# Carregando a planilha
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
gc = gspread.service_account(filename="Inserir o caminho da API.json")
planilha = gc.open_by_key("Inserir a chave da planilha")
worksheet = planilha.worksheet("Inserir o nome da aba")
valores = worksheet.get_all_values()
print(valores)

mensagem = f"""<p>Prezado (a)</p>
        <p> Favor analisar a viagem reforço anexada a este e-mail.</p>
        <p> Em caso de validação, favor assinar e incluir na pasta do seguinte link: "https://drive.google.com/drive/folders/1PRb3Ht8U-3CC1gFyOe8Y0Lb8JclL9z_k?usp=drive_link".</p>
        <p>Atenciosamente,</p>: 
"""

tabela_gerencias = pd.DataFrame(
    {
        "Gerencia": ["Inserir lista das gerências"],
        "Email": ["Inserir Linha de E-mails"
        ],
    }
)

for i, row in enumerate(valores):
    if i > 0:
        carimbo = row[0]
        email = row[1]
        motivo = row[2]
        data = row[3]
        sistema = row[4]
        linha = row[5]
        sublinha = row[6]
        partidasPC = row[7]
        NºdeOrdem = row[8]
        Placa = row[9]
        CodViagem = row[10]
        HorarioAberturaMCO = row[11]
        Agente = row[12]
        BT = row[13]
        Gerencia = row[14]
        
        filtro = tabela_gerencias.loc[tabela_gerencias['Gerencia'] == Gerencia]
        destinatario = filtro.iloc[0]["Email"]
        data_formatada = datetime.strptime(carimbo, "%d/%m/%Y %H:%M:%S")
        # Verificar a quantidade atual de linhas preenchidas na planilha
        if data_formatada.date() == (datetime.today().date() - timedelta(days=0)):
            print(destinatario)
            #     # Chamar a função para enviar o e-mail
            pdf = cria_PDF(i)
            enviar_email("Viagem Reforço", destinatario, mensagem, pdf)
