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

df_existing = df_existing.drop_duplicates(subset=["Venda"], keep="last")

for sale in sales_data:
    if "payment" in sale and "installments" in sale["payment"]:
        
        customer = sale.get("customer", {})
        seller = sale.get("seller", {})
        payment = sale.get("payment", {})
        emission = sale.get("emission", "").split("T")[0]
        installments = sale["payment"].get("installments", [])
        filter_data = datetime.strptime(emission, "%Y-%m-%d")
        data = datetime.strptime(emission, "%Y-%m-%d").strftime("%d/%m/%Y")
        metodo = payment.get("method")
        if metodo == "CASH":
            continue

        if sale.get("seller", {}).get("name") == 'Financeiro':
            continue
        
        number = sale.get("number")
        
        num_installments = len(installments)

        max_installment_num = 0
        for installment in installments:
            num = installment.get("number", 1)
            max_installment_num = max(max_installment_num, num)
        
        parcel_values = [None] * 10
        status_colors = [None] * 10
        font_colors = [None] * 10
        
        
        for installment in installments:
            index = installment.get("number", 1) - 1
            if index < 10:
                parcel_values[index] = installment.get("value")
                
                status = installment.get("status", "").upper()
                due_date_str = installment.get("due_date", "").split("T")[0]
                
                if due_date_str:
                    due_date = datetime.strptime(due_date_str, "%Y-%m-%d").date()
                    is_late = due_date < today and status == "PENDING"
                else:
                    is_late = False
                
                if status == "ACQUITTED":
                    status_colors[index] = (0.0, 1.0, 0.0)  
                    font_colors[index] = (0.0, 0.0, 0.0)
                elif is_late:
                    status_colors[index] = (1.0, 0.0, 0.1)  
                    font_colors[index] = (1.0, 1.0, 1.0)
                else:
                    status_colors[index] = (1.0, 1.0, 0.8)
                    font_colors[index] = (0.0, 0.0, 0.0)
        
        
        existing_row_idx = df_existing.index[df_existing.get("Venda") == number].tolist()
        
        if not existing_row_idx and filter_data.date().month == 2:
                
            filtered_installments.append([
                seller.get("name"),
                customer.get("name"),
                sale.get("number"),
                data,
                sale.get("total"),
                payment.get("method").replace("BANKING_BILLET", f"Boleto {num}x").replace("OTHER", "Cartão"),
                *parcel_values
            ])
            cell_formats.append((status_colors,font_colors))

def parse_date(data):
    try:
        return datetime.strptime(data, "%d/%m/%Y").date()
    except ValueError:
        return datetime.min

combined_data = list(zip(filtered_installments, cell_formats))

combined_data.sort(key=lambda x: (
    x[0][0].strip().lower() if x[0][0] else "",
    parse_date(x[0][3])
))

if existing_data:
    filtered_installments, cell_formats = zip(*combined_data)
    filtered_installments = list(filtered_installments)
    cell_formats = list(cell_formats)
else:
    pass


ultima_linha = len(df_existing) + 9
if ultima_linha < 9:
    ultima_linha = 9


num_linhas = len(filtered_installments)
range_inicio = ultima_linha
range_fim = range_inicio + num_linhas - 1
intervalo = f"A{range_inicio}:P{range_fim}"

sheet.update(filtered_installments, intervalo)

requests = []

for i in range(10):
    for row_offset, (colors, font_colors) in enumerate(cell_formats):
        row = range_inicio + row_offset  # Define a linha correta

        color = colors[i] if i < len(colors) else None
        font_color = font_colors[i] if i < len(font_colors) else (0, 0, 0)

        if color:
            requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": sheet.id,
                        "startRowIndex": row - 1,
                        "endRowIndex": row,
                        "startColumnIndex": 6 + i,
                        "endColumnIndex": 7 + i
                    },
                    "rows": [{
                        "values": [{
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": color[0],
                                    "green": color[1],
                                    "blue": color[2]
                                },
                                "textFormat": {
                                    "foregroundColor": {
                                        "red": font_color[0],
                                        "green": font_color[1],
                                        "blue": font_color[2]
                                    }
                                }
                            }
                        }]
                    }],
                    "fields": "userEnteredFormat(backgroundColor,textFormat.foregroundColor)"
                }
            })

print(json.dumps(requests, indent=2))
if requests:
    sheet.spreadsheet.batch_update({"requests": requests})  # Fixed batch_update call