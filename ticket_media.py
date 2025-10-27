import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.image import MIMEImage

# -------------------
# Importar base de dados
tabela_val = pd.read_excel("Vendas.xlsx")# Sua tabela
pd.set_option("display.max_columns", None)

# -------------------
# Faturamento por loja
faturamento = tabela_val.groupby("ID Loja")[["Valor Final"]].sum().reset_index()

faturamento['Valor Final'] = faturamento['Valor Final'].astype(float)

# Vendas por loja
vendas = tabela_val.groupby("ID Loja")[["Quantidade"]].sum().reset_index()

# -------------------
# Ticket médio
ticket = faturamento.merge(vendas, on="ID Loja")
ticket["Ticket Médio"] = ticket["Valor Final"] / ticket["Quantidade"]

# Formatação 
faturamento_formatado = faturamento.copy()
faturamento_formatado['Valor Final'] = faturamento_formatado['Valor Final']\
    .apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

ticket_format = ticket.copy()
ticket_format['Ticket Médio'] = ticket_format['Ticket Médio']\
    .apply(lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))

# -------------------
# Função para enviar o email
def enviar_email():
    msg = MIMEMultipart()
    msg['Subject'] = "Relatório de Vendas" # Assunto
    msg['From'] = 'botmailsender4@gmail.com' # Remetente
    msg['To'] = 'matheusfs589@gmail.com' # Destinatário
    senha = 'zxci xvhs pkqp xiwk'  # Senha de App (Senha normal não funciona)


    corpo_email = f"""
    <html>
    <head>
        <style>
            table {{ border-collapse: collapse; width: 80%; margin: 0 auto; }}
            th, td {{ border: 1px solid #ddd; padding: 8px; text-align: center; }}
            th {{ background-color: #f2f2f2; }}
            h2 {{ text-align: center; color: #333; }}
            p {{ text-align: center; font-family: Arial, sans-serif; }}
        </style>
    </head>
    <body>
        <p>Olá,</p>
        <p>Segue o Relatório de Vendas por Loja:</p>

        <h2>Faturamento</h2>
        {faturamento_formatado.to_html(index=False, border=0, justify='center')}

        <h2>Vendas</h2>
        {vendas.to_html(index=False, border=0, justify='center')}

        <h2>Ticket Médio</h2>
        {ticket_format[["ID Loja","Ticket Médio"]].to_html(index=False, border=0, justify='center')}

        <p>Atenciosamente,<br>
        Matheus F.<br>
        <img src="cid:logo" width="450" height="auto"></p>
    </body>
    </html>
    """

    msg.attach(MIMEText(corpo_email, 'html'))

        # imagem assinatura de email
    with open('c:/Users/daniv/Downloads/nome.png', 'rb') as f:
        img = MIMEImage(f.read())
        img.add_header('Content-ID', '<logo>')
        msg.attach(img)

    # Enviar email
    with smtplib.SMTP('smtp.gmail.com', 587) as s:
        s.starttls()
        s.login(msg['From'], senha)
        s.send_message(msg)

    print('Email enviado com sucesso!')



enviar_email()
