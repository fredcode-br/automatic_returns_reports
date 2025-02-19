import win32com.client
import os
import smtplib
import logging
import sys
import time
from win32com.client import constants as xl
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders


import sys
import logging

# Configuração de logging
logging.basicConfig(
    filename="relatorios.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%d/%m/%Y %H:%M:%S"
)

# Criar uma classe para duplicar a saída do print
class DualStream:
    def write(self, message):
        logging.info(message)  # Registra no log (mantém linhas em branco)
        sys.__stdout__.write(message)  # Exibe no terminal

    def flush(self):
        sys.__stdout__.flush()

# Redirecionar stdout para DualStream
sys.stdout = DualStream()

def enviar_email(destinatario, assunto, corpo, arquivos_anexos):
    try:
        remetente = "relatorios@bioleve.com.br"
        senha = "M1nFl@145236"
        msg = MIMEMultipart()
        msg['From'] = remetente
        msg['To'] = destinatario
        msg['Subject'] = assunto

        msg.attach(MIMEText(corpo, 'plain'))

        print(f"Enviando e-mail para {destinatario}...")

        # Anexar os arquivos PDF
        for arquivo in arquivos_anexos:
            if os.path.exists(arquivo):
                with open(arquivo, "rb") as anexo:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(anexo.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(arquivo)}")
                    msg.attach(part)
            else:
                print(f"Aviso: Arquivo não encontrado - {arquivo}")


        with smtplib.SMTP_SSL('smtp.mailcorp.com.br', 465) as servidor:
            servidor.login(remetente, senha)
            servidor.sendmail(remetente, destinatario, msg.as_string())
        print(f"E-mail enviado para {destinatario}.")
    except Exception as e:
        print(f"Erro ao enviar o e-mail para {destinatario}: {e}")

def relatorios(workbook, planilha_relatorio, setor, pasta_destino):
    try:
        os.makedirs(pasta_destino, exist_ok=True)  # Garante que a pasta de destino exista
        sheet = workbook.Sheets(planilha_relatorio)

        caminho_pdf = os.path.join(pasta_destino, f'{planilha_relatorio}_{setor}.pdf')
        sheet.ExportAsFixedFormat(0, caminho_pdf)  # 0 = PDF

        return caminho_pdf
    
    except Exception as e:
        print(f"\nErro ao gerar o relatório: {e}")
        return None
    
    finally:
        "Relatório criádo com sucesso!"

def atualizarDados(caminho_arquivo_xlsm, data_inicial, data_final, pasta_destino):
    try:
        # Fechar instâncias do Excel antes de abrir um novo arquivo
        fechar_instancias_excel()

        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Mantenha o Excel invisível para evitar interferências

        print("Abrindo o arquivo...")
        workbook = excel.Workbooks.Open(os.path.abspath(caminho_arquivo_xlsm))
        print("Atualizando dados...")
        # workbook.RefreshAll()
        # time.sleep(1) # Tempo para atualizar os dados
        print("Salvando o arquivo...")
        workbook.Save()
        # time.sleep(2)  # Garantir que o Excel processe o comando Save

        if not os.path.exists(caminho_arquivo_xlsm):
            print(f"Erro: O arquivo {caminho_arquivo_xlsm} não foi salvo corretamente.")
        else:
            print(f"Arquivo salvo com sucesso em {caminho_arquivo_xlsm}.")

        sheet = workbook.Sheets("Ocorrências")

        data_inicial = data_inicial.strftime("%d/%m/%Y")
        data_final = data_final.strftime("%d/%m/%Y")

        tabela = sheet.ListObjects("Tabela_Ocorrências")

        coluna_data = 2       # Coluna B = DATA
        coluna_local = 13     # Coluna M = EMISSÃO NF
        coluna_setor = 14     # Coluna N = SETOR RESP.

        print("Limpando filtros anteriores...")
        tabela = sheet.ListObjects("Tabela_Ocorrências")
        if tabela.ShowAutoFilter:
            tabela.AutoFilter.ShowAllData()  # Remove todos os filtros antes de aplicar novos

        print("Aplicando filtro na coluna DATA...")
        # Aplicar filtro de data
        tabela.Range.AutoFilter(
            Field=int(coluna_data),
            Criteria1=">=" + str(data_inicial),
            Operator=1,
            Criteria2="<=" + str(data_final)
        )

        setores = ['COM','CQ','FIS','EXP','LOG']
        locais = ["LINDÓIA", "SÃO BERNARDO"]

        for setor in setores:
            tabela.Range.AutoFilter(
                Field=int(coluna_setor),
                Criteria1=setor
            )
            
            if setor == "EXP" or setor == "LOG":
                for local in locais:
                    print(f" - Filtrando pelo local: {local}... \n")
                    tabela.Range.AutoFilter(
                        Field=int(coluna_local),
                        Criteria1=local
                    )
                    relatorio  = relatorios(workbook, "Ocorrências", (setor+"_"+local), pasta_destino)
                    email_destino = pegar_email(workbook, setor, local) 
            else:
                print(f"\n***** GERANDO E-MAILS PARA O SETOR {setor} *****\n")
                relatorio  = relatorios(workbook, "Ocorrências", setor, pasta_destino)
                email_destino = pegar_email(workbook, setor) 

            if not email_destino:
                print(f"⚠️ Atenção: Nenhum e-mail encontrado para o setor {setor}.")
                continue  # Pula para o próximo setor


            assunto = f"OCORRÊNCIAS DEVOLUÇÕES"
            corpo = (
                f"Segue a relação de ocorrências referente ao seu setor, da semana passada.\n\n"
                "Favor não responder a este e-mail.\n\n"
                "Atenciosamente,\nEquipe TI Bioleve"
            )
            
            # enviar_email(dest["email"], assunto, corpo, [relatorio_em_aberto, relatorio_faturado])
    finally:
        workbook.Close(SaveChanges=True)
        excel.Quit()

    print("Processo concluído!")

def pegar_email(workbook, setor, local = ""):
    print(local)
    sheetEmails = workbook.Sheets("Emails")
    tabelaEmails = sheetEmails.ListObjects("Tabela_Email_Setor")

    # Remove todos os filtros antes de aplicar novos
    if tabelaEmails.ShowAutoFilter:
        tabelaEmails.AutoFilter.ShowAllData()  

    # --- Filtrando o e-mail na Tabela_Email_Setor ---
    print(f"Buscando e-mail para o setor {setor}...")
    tabelaEmails.Range.AutoFilter(
        Field=2,  # Coluna "Setor"
        Criteria1=setor
    )

    # Pegando a primeira linha visível após o filtro
    linhas_visiveis = tabelaEmails.DataBodyRange.SpecialCells(12)
    # Pegando o valor da terceira coluna (Email) na primeira linha visível
    email_destino = linhas_visiveis.Cells(1, 3).Value if linhas_visiveis.Rows.Count > 0 else None
        
    print(email_destino) 

    return email_destino

def fechar_instancias_excel():
    os.system("taskkill /f /im excel.exe >nul 2>&1")

def enviar_logs_do_dia(destinatario):
    try:
        hoje = datetime.now().strftime("%d/%m/%Y")
        logs_do_dia = []

        # Ler os logs do arquivo original
        with open("relatorios.log", "r") as log_file:
            for linha in log_file:
                if linha.startswith(hoje):  # Filtrar apenas as linhas do dia atual
                    logs_do_dia.append(linha)

        if not logs_do_dia:
            print("Nenhum log do dia atual encontrado.")
            return

        # Criar um arquivo temporário com os logs do dia
        caminho_temporario = "logs_do_dia.log"
        with open(caminho_temporario, "w") as temp_file:
            temp_file.writelines(logs_do_dia)

        # Enviar o arquivo por e-mail
        assunto = "Logs Pedidos em Aberto e Faturados"
        corpo = "Segue em anexo os logs gerados no dia atual."
        caminho_temporario = [f'C:\\Scripts\\Relatórios_Devoluções\\{caminho_temporario}']

        enviar_email(destinatario, assunto, corpo, caminho_temporario)

        print("Logs do dia enviados.")
    except Exception as e:
        print(f"Erro ao enviar os logs do dia: {e}")

# Caminhos dos arquivos
caminho_arquivo_xlsm = r"C:\Scripts\Relatórios_Devoluções\Dados.xlsm"
pasta_destino = r"C:\Scripts\Relatórios_Devoluções\relatorios"

data_final = datetime.now()
data_inicial = datetime(data_final.year, data_final.month, (data_final.day - 7))

# Execução principal
atualizarDados(caminho_arquivo_xlsm, data_inicial, data_final, pasta_destino)

# Enviar logs do dia
# enviar_logs_do_dia("relatorios@bioleve.com.br")