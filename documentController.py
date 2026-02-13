from google.oauth2 import service_account
from googleapiclient.discovery import build
import pandas as pd
import os
from docx import Document
from datetime import datetime

OUTPUT = "/home/wal/AutoDoc2.0/output"
INPUT = "/home/wal/AutoDoc2.0/input"


def safe_get(row, column):
    """Evita erro se coluna não existir ou estiver vazia"""
    return str(row[column]) if column in row and pd.notna(row[column]) else ""


def start_doc_process_from_row(row):
    # Processar data com segurança
    data_str = safe_get(row, "Data da Defesa")

    numero_dia = numero_mes = numero_ano = ""

    if data_str:
        try:
            data_obj = datetime.strptime(data_str, "%d/%m/%Y")
            numero_dia = data_obj.strftime("%d")
            numero_mes = data_obj.strftime("%B")
            numero_ano = data_obj.strftime("%Y")
        except ValueError:
            print(f"Data inválida: {data_str}")

    word_dict = {
        "destino": safe_get(row, "Email Responsável"),
        "nome_coordenador": "Leandro Colombi",
        "numero_dia": numero_dia,
        "nome_mes": numero_mes,
        "numero_ano": numero_ano,
        "numero_hora": safe_get(row, "Horário da Defesa"),
        "numero_sala": safe_get(row, "Sala da Defesa"),
        "link_sala": safe_get(row, "Se online, link da sala"),
        "nome_completo_aluno": safe_get(row, "Nome Completo do Aluno"),
        "titulo_tese": safe_get(row, "Título Tese"),
        "nome_orientador1": safe_get(row, "Nome Orientador"),
        "nome_orientador2": safe_get(row, "Nome Coorientador"),
        "nome_membro_interno": safe_get(row, "Nome Membro Interno"),
        "nome_membro_externo": safe_get(row, "Nome Membro Externo")
    }

    print("Dados Coletados:", word_dict)
    return word_dict

def process_documents(word_dict):
    """
    Processa todos os arquivos .docx da pasta de entrada,
    substitui palavras conforme o dicionário recebido e
    salva os documentos na pasta do aluno dentro de /output.
    """

    os.makedirs(OUTPUT, exist_ok=True)

    #Percorre os arquivos da pasta de Entrada
    for filename in os.listdir(INPUT):

        if filename.endswith(".docx"):

            input_path = os.path.join(INPUT, filename)
            output_path = os.path.join(OUTPUT, filename)

            doc = Document(input_path)
            replace_words_in_document(doc, word_dict)

            if filename == "Modelo_Ata_de_Defesa.docx":

                # APROVADO
                replace_words_in_document(doc, {"resultado": "APROVADO"})
                approved_path = os.path.join(
                    OUTPUT,
                    f"{os.path.splitext(filename)[0]}_APROVADO.docx"
                )
                doc.save(approved_path)

                # REPROVADO
                doc = Document(input_path)
                replace_words_in_document(doc, word_dict)
                replace_words_in_document(doc, {"resultado": "REPROVADO"})
                reproved_path = os.path.join(
                    OUTPUT,
                    f"{os.path.splitext(filename)[0]}_REPROVADO.docx"
                )
                doc.save(reproved_path)

            else:
                doc.save(output_path)

    return list_doc(OUTPUT)


def replace_words_in_document(doc, word_dict):
    for para in doc.paragraphs:
        for run in para.runs:
            for old_word, new_word in word_dict.items():
                if old_word in run.text:
                    run.text = run.text.replace(old_word, str(new_word))


def list_doc(caminho_pasta):
    """
    Retorna uma lista com o caminho completo
    de todos os arquivos da pasta informada.
    """
    return [
        os.path.join(caminho_pasta, arquivo)
        for arquivo in os.listdir(caminho_pasta)
        if os.path.isfile(os.path.join(caminho_pasta, arquivo))
    ]
