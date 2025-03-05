import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import unicodedata
import json
import pandas as pd
import time

def setup_google_sheets():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("src/common/config/api-planilha-rodrigo.json", scopes=scope)
    client = gspread.authorize(creds)
    
    SHEET_NAME = "Comissão Time de Vendas TESTE"
    WORKSHEET_NAME = "Carlos Louback"
    return client.open(SHEET_NAME).worksheet(WORKSHEET_NAME)

def get_existing_data(sheet):
    existing_data = sheet.get_all_values()[7:]
    if existing_data:
        df_existing = pd.DataFrame(existing_data[1:], columns=existing_data[0])
    else:
        colunas = ["Representante Comercial", "Cliente", "Venda", "Data", "Valor Total", "Forma de Pagamento"] + [f"Parcela {i+1}" for i in range(10)]
        df_existing = pd.DataFrame(columns=colunas)

    if "Venda" in df_existing.columns:
        df_existing["Venda"] = pd.to_numeric(df_existing["Venda"], errors="coerce")
    
    return df_existing

def load_sales_data():
    with open('src/common/input_files/data.json', "r", encoding="utf-8") as file:
        return json.load(file)

def process_installments(sale, today):
    if "payment" not in sale or "installments" not in sale["payment"]:
        return None

    customer = sale.get("customer", {})
    seller = sale.get("seller", {})
    payment = sale.get("payment", {})
    emission = sale.get("emission", "").split("T")[0]
    installments = sale["payment"].get("installments", [])
    filter_data = datetime.strptime(emission, "%Y-%m-%d")
    data = datetime.strptime(emission, "%Y-%m-%d").strftime("%d/%m/%Y")
    metodo = payment.get("method")
    
    if metodo == "CASH" or sale.get("seller", {}).get("name") == 'Financeiro':
        return None

    number = sale.get("number")
    num_installments = len(installments)
    max_installment_num = max((installment.get("number", 1) for installment in installments), default=0)
    
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
                status_colors[index] = (0.0, 255/255, 0.0)
                font_colors[index] = (0.0, 0.0, 0.0)
            elif is_late:
                status_colors[index] = (255/255, 0.0, 0.0)
                font_colors[index] = (1.0, 1.0, 1.0)
            else:
                status_colors[index] = (255/255, 242/255, 204/255)
                font_colors[index] = (0.0, 0.0, 0.0)
    
    return {
        'seller': seller.get("name"),
        'customer': customer.get("name"),
        'number': number,
        'data': data,
        'total': sale.get("total"),
        'payment_method': payment.get("method").replace("BANKING_BILLET", f"Boleto {num_installments}x").replace("OTHER", "Cartão"),
        'parcel_values': parcel_values,
        'status_colors': status_colors,
        'font_colors': font_colors,
        'filter_data': filter_data
    }

def update_sheet(sheet, filtered_installments, cell_formats, range_inicio):
    num_linhas = len(filtered_installments)
    range_fim = range_inicio + num_linhas - 1
    intervalo = f"A{range_inicio}:P{range_fim}"
    sheet.update(filtered_installments, intervalo)

    requests = []
    for i in range(10):
        for row_offset, (colors, font_colors) in enumerate(cell_formats):
            row = range_inicio + row_offset
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

    if requests:
        sheet.spreadsheet.batch_update({"requests": requests})

def main_insertData():
    sheet = setup_google_sheets()
    df_existing = get_existing_data(sheet)
    sales_data = load_sales_data()
    
    today = datetime.today().date()
    filtered_installments = []
    cell_formats = []
    
    df_existing = df_existing.drop_duplicates(subset=["Venda"], keep="last")
    
    for sale in sales_data:
        installment_data = process_installments(sale, today)
        if installment_data:
            existing_row_idx = df_existing.index[df_existing.get("Venda") == installment_data['number']].tolist()
            
            if not existing_row_idx and installment_data['filter_data'].date().month == 2:
                filtered_installments.append([
                    installment_data['seller'],
                    installment_data['customer'],
                    installment_data['number'],
                    installment_data['data'],
                    installment_data['total'],
                    installment_data['payment_method'],
                    *installment_data['parcel_values']
                ])
                cell_formats.append((installment_data['status_colors'], installment_data['font_colors']))

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

    if filtered_installments:
        filtered_installments, cell_formats = zip(*combined_data)
        filtered_installments = list(filtered_installments)
        cell_formats = list(cell_formats)

        ultima_linha = len(df_existing) + 9
        if ultima_linha < 9:
            ultima_linha = 9

        update_sheet(sheet, filtered_installments, cell_formats, ultima_linha)

if __name__ == "__main__":
    main_insertData()