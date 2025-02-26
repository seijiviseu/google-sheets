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
df_existing = pd.DataFrame(existing_data[1:], columns=existing_data[0]) if len(existing_data) > 1 else pd.DataFrame()

def remove_acentos(input_str):
    if input_str is None:
        return None
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([char for char in nfkd_form if not unicodedata.combining(char)])

raw_data_path = 'src/raw_data/data.json'

with open(raw_data_path, "r", encoding="utf-8") as file:
    sales_data = json.load(file)

filtered_installments = []
cell_formats = []  

today = datetime.today().date()  

for sale in sales_data:
    if "payment" in sale and "installments" in sale["payment"]:
        
        customer = sale.get("customer", {})
        seller = sale.get("seller", {})
        payment = sale.get("payment", {})
        number = sale.get("number")
        emission = sale.get("emission", "").split("T")[0]
        installments = sale["payment"].get("installments", [])
        data = datetime.strptime(emission, "%Y-%m-%d").strftime("%d/%m/%Y")
        metodo = payment.get("method")
        if metodo == "CASH":
            continue

        parcel_values = [None] * 10
        status_colors = [None] * 10
        font_colors = [None] * 10

        if "Venda" not in df_existing.columns:
            raise KeyError("A coluna 'Venda' não foi encontrada. Verifique o cabeçalho da planilha.")
        df_existing["Venda"] = pd.to_numeric(df_existing["Venda"], errors="coerce")
    else:
        colunas = ["Representante Comercial", "Cliente", "Venda", "Data", "Valor Total", "Forma de Pagamento"] + [f"Parcela {i+1}" for i in range(10)]
        df_existing = pd.DataFrame(columns=colunas)

    batch_updates = []  # Para atualizações em massa
    new_rows = []  # Para novas linhas a serem adicionadas

    for row in filtered_installments:
        if isinstance(row, dict) and "installments" in row and isinstance(row["installments"], list):
            for installment in row["installments"]:
                index = installment.get("number", 1)
                status = installment.get("status", "")
                due_date_str = installment.get("due_date", "")
                sale_number = row[2]
                existing_row_idx = df_existing.index[df_existing.get("Venda") == sale_number].tolist()

                if due_date_str:
                    due_date = datetime.strptime(due_date_str, "%Y-%m-%d").date()
                    is_late = due_date < today and status == "PENDING"
                else:
                    is_late = False

                if existing_row_idx:  # Atualiza se já existe
                    row_idx = existing_row_idx[0] + 7 + 2
                    cell_range = f"A{row_idx}:P{row_idx}"
                    batch_updates.append({"range": cell_range, "values": [row]})
                else:  # Adiciona se não existir
                    new_rows.append(row)

    # **Faz todas as atualizações de uma vez**
                if batch_updates:
                    sheet.batch_update(batch_updates)
                    print(f"✅ {len(batch_updates)} vendas atualizadas!")

                # **Adiciona novas linhas de uma vez**
                if new_rows:
                    sheet.append_rows(new_rows, value_input_option="RAW")
                    print(f"➕ {len(new_rows)} novas vendas adicionadas!")

                # Definir cores com base no status
                if status == "ACQUITTED":
                    status_colors[index] = (0.0, 1.0, 0.0)  # Verde
                    font_colors[index] = (0, 0, 0)
                elif is_late:
                    status_colors[index] = (1, 0.0, 0.1)  # Vermelho
                else:
                    status_colors[index] = (1, 1, 0.8)  # Amarelo
                    font_colors[index] = (0, 0, 0)

    filtered_installments.append([
        seller.get("name"),
        customer.get("name"),
        sale.get("number"),
        data,
        sale.get("total"),
        payment.get("method").replace("BANKING_BILLET", f"Boleto {len(installments)}x").replace("OTHER", "Cartão"),
        *parcel_values
    ])
    cell_formats.append((status_colors))

ultima_linha = len(df_existing) + 1
if ultima_linha < 9:
    ultima_linha = 9

range_inicio = ultima_linha
num_linhas = len(filtered_installments)
range_fim = range_inicio + num_linhas - 1
intervalo = f"A{range_inicio}:P{range_fim}"  

batch_updates = []

for row in filtered_installments:
    sale_number = row[2]
    existing_row_idx = df_existing.index[df_existing.get("Venda") == sale_number].tolist()

    if existing_row_idx:
        batch_updates.append({"range": intervalo, "values": [row]})
    else:
        new_rows.append(row)

# 3️⃣ 🚀 **Envia todas as atualizações em uma única requisição**
if batch_updates:
    sheet.batch_update(batch_updates)
    print(f"✅ {len(batch_updates)} vendas atualizadas!")

if new_rows:
    sheet.append_rows(new_rows, value_input_option="RAW")
    print(f"➕ {len(new_rows)} novas vendas adicionadas!")

requests = []
for i, (colors) in enumerate(cell_formats):
    row = range_inicio + i  
    for j, color in enumerate(colors):
        if color:
            requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": sheet.id,
                        "startRowIndex": row - 1,
                        "endRowIndex": row,
                        "startColumnIndex": 6 + j,  
                        "endColumnIndex": 7 + j
                    },
                    "rows": [{
                        "values": [{
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": color[0],
                                    "green": color[1],
                                    "blue": color[2]
                                },
                            }
                        }]
                    }],
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })


if requests:
    body = {"requests": requests}
    sheet.spreadsheet.batch_update(body)

print("Dados inseridos e cores aplicadas com sucesso!")
time.sleep(1)
