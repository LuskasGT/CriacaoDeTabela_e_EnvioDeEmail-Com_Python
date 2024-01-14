import os
import pandas as pd
import smtplib
import email.message
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime

# Achando os arquivos
caminho = r"C:\Users\lucas\OneDrive\Documentos\GitHub\bases"
arquivos = os.listdir(caminho)
print(arquivos)

tabela_consolidada = pd.DataFrame()

# Criando o xlsx e organizando o arquivo
for nome_arquivo in arquivos:
    arquivo_path = os.path.join(caminho, nome_arquivo)
    try:
        tabela_vendas = pd.read_csv(arquivo_path, encoding='utf-8')  # Especifique a codificação
        tabela_vendas["Data de Venda"] = pd.to_datetime("01/01/1900") + pd.to_timedelta(tabela_vendas["Data de Venda"], unit="d")
        tabela_consolidada = pd.concat([tabela_consolidada, tabela_vendas])
    except pd.errors.ParserError as e:
        print(f"Erro ao ler o arquivo {arquivo_path}: {e}")
    except UnicodeDecodeError:
        print(f"Erro de codificação ao ler o arquivo {arquivo_path}")

tabela_consolidada = tabela_consolidada.sort_values(by="Data de Venda").reset_index(drop=True)#o reset index é parann ficar com numeros repetidos por conta dajunção das tabelas

# Salvando o arquivo Vendas.xlsx
data_hoje = datetime.today().strftime("%d-%m-%Y")
arquivo_excel = os.path.join(caminho, f"Vendas_{data_hoje}.xlsx")
tabela_consolidada.to_excel(arquivo_excel, index=False)

# Criando o envio de email
def send_email():
    corpo_email = f"""
    <p>Olá Lucas,</p>
    <br>
    <p>Segue em anexo o relatório de Vendas atualizado do dia {data_hoje}.</p>
    <br>
    <p>Obrigado</p>
    """

    msg = MIMEMultipart()
    msg['Subject'] = f"Relatorio de Vendas do dia {data_hoje}"
    msg['From'] = "lucas.g.theodooro@gmail.com"
    msg['To'] = "lucas.g.theodooro@gmail.com"
    password = 'SENHA AQUI(NN A DO SEU EMAIL)'
    
    # Adicionando corpo do email
    msg.attach(MIMEText(corpo_email, 'html'))

    # Adicionando anexo
    anexo_path = arquivo_excel
    anexo_nome = f"Vendas_{data_hoje}.xlsx"
    attachment = open(anexo_path, 'rb')

    part = MIMEBase('application', 'octet-stream')
    part.set_payload((attachment).read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {anexo_nome}")

    msg.attach(part)

    # Configurar servidor SMTP
    s = smtplib.SMTP('smtp.gmail.com:587')
    s.starttls()

    # Login com credenciais para o envio de email
    s.login(msg['From'], password)

    # Enviar email
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))

    # Fechar a conexão com o servidor
    s.quit()
    print('Email enviado com anexo')

send_email()
