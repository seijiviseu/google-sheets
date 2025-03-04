import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import unicodedata
import json
import pandas as pd
import time

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file("src/token/api-planilha-rodrigo.json", scopes=scope)
client = gspread.authorize(creds)


SHEET_NAME = "Comissão Time de Vendas TESTE"
WORKSHEET_NAME = "Carlos Louback"
sheet = client.open(SHEET_NAME).worksheet(WORKSHEET_NAME)

existing_data = sheet.get_all_values()[7:]
if existing_data:
    df_existing = pd.DataFrame(existing_data[1:], columns=existing_data[0])
else:
    colunas = ["Representante Comercial", "Cliente", "Venda", "Data", "Valor Total", "Forma de Pagamento"] + [f"Parcela {i+1}" for i in range(10)]
    df_existing = pd.DataFrame(columns=colunas)

if "Venda" in df_existing.columns:
    df_existing["Venda"] = pd.to_numeric(df_existing["Venda"], errors="coerce")

raw_data_path = 'src/raw_data/delete_data.csv'
df_delete = pd.read_csv(raw_data_path)

if "Venda" in df_delete.columns:
    df_delete["Venda"] = pd.to_numeric(df_delete["Venda"], errors="coerce")

# 🔹 Identificar as linhas a serem deletadas
for venda in df_delete["Venda"]:
    index_list = df_existing[df_existing["Venda"] == venda].index.tolist()

    # Deletar cada linha encontrada
    for idx in sorted(index_list, reverse=True):  # Ordena para deletar de baixo para cima
        sheet.delete_rows(idx + 9) # Deletando a venda 17793
