from google.oauth2 import service_account
from googleapiclient.discovery import build
import pandas as pd
import os
import shutil
from docx import Document
from datetime import datetime
from documentController import safe_get, start_doc_process_from_row, process_documents
from emailController import EmailSender, create_email
import time

# =========================
# CONFIGURAÇÕES GOOGLE
# =========================

SERVICE_ACCOUNT_FILE = 'key.json'

SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

SPREADSHEET_ID = '1FU4nVHzOupO8fhV1jqA7JswZoOE8OfJ444bNb7jZFCo'
RANGE_NAME = 'dados'

COORDENADOR = "walcandeia@gmail.com"

OUTPUT_PATH = "/home/wal/AutoDoc2.0/output"
# =========================
# EXECUÇÃO PRINCIPAL
# =========================

def main():
    sender = EmailSender()
    sender.connect()

    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )

    service = build('sheets', 'v4', credentials=creds)

    result = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=RANGE_NAME
    ).execute()

    values = result.get('values', [])

    if not values:
        print("Nenhum dado encontrado.")
        return

    headers = values[0]
    data = values[1:]

    if "Status" not in headers:
        headers.append("Status")
        for row in data:
            row.append("")

    normalized_data = [
        row + [""] * (len(headers) - len(row))
        if len(row) < len(headers)
        else row[:len(headers)]
        for row in data
    ]

    df = pd.DataFrame(normalized_data, columns=headers)
    status_index = headers.index("Status")

    for idx, row in df.iterrows():

        if safe_get(row, "Status").strip().upper() == "PROCESSADO":
            continue  # já processado

        print(f"Processando: {safe_get(row, 'Nome Completo do Aluno')}")

        word_dict = start_doc_process_from_row(row)
        attachments = process_documents(word_dict)

        email_data = create_email(word_dict)
        email_data.attachments = attachments

        sender.send_email(email_data, COORDENADOR)
        sender.send_email(email_data, word_dict.get("destino"))

        # Atualiza status na planilha
        sheet_row_number = idx + 2  # +2 porque:
        # 1 = cabeçalho
        # 1 = index começa em 0

        update_range = f"{RANGE_NAME}!{chr(65 + status_index)}{sheet_row_number}"

        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=update_range,
            valueInputOption="RAW",
            body={
                "values": [["PROCESSADO"]]
            }
        ).execute()

        print("✔ Marcado como PROCESSADO\n")

        if not os.path.isdir(OUTPUT_PATH):
            print("Pasta não encontrada.")
            return False

        for item in os.listdir(OUTPUT_PATH):
            item_path = os.path.join(OUTPUT_PATH, item)

            try:
                if os.path.isfile(item_path) or os.path.islink(item_path):
                    os.remove(item_path)  # remove arquivo
                elif os.path.isdir(item_path):
                    shutil.rmtree(item_path)  # remove subpasta inteira
            except Exception as e:
                print(f"Erro ao remover {item_path}: {e}")

        print("Pasta esvaziada com sucesso.")
        return True

if __name__ == "__main__":
    while True:
        try:
            print("Verificando Formulários...")
            main()
        except Exception as e:
            print("Erro:", e)
        
        time.sleep(60)

