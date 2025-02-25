import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import unicodedata
import json
import pandas as pd

# Função para remover acentos
def remove_acentos(input_str):
    if input_str is None:
        return None
    nfkd_form = unicodedata.normalize('NFKD', input_str)
    return "".join([char for char in nfkd_form if not unicodedata.combining(char)])

# Caminho do arquivo JSON
raw_data_path = 'raw_data/data.json'

# Carregar os dados do JSON
with open(raw_data_path, "r", encoding="utf-8") as file:
    sales_data = json.load(file)

# Filtrar os dados relevantes
filtered_installments = []
cell_formats = []  # Lista para armazenar formatações das células

today = datetime.today().date()  # Data atual

for sale in sales_data:
    if "payment" in sale and "installments" in sale["payment"]:
        customer = sale.get("customer", {})
        seller = sale.get("seller", {})
        payment = sale.get("payment", {})
        emission = sale.get("emission", "").split("T")[0]
        installments = sale["payment"].get("installments", [])
        data = datetime.strptime(emission, "%Y-%m-%d").strftime("%d/%m/%Y")

        # Criar placeholders para parcelas
        parcel_values = [None] * 10
        status_colors = [None] * 10  # Lista para armazenar cores das células
        font_colors = [None] * 10  # Lista para armazenar cores da fonte

        for installment in installments:
            index = installment.get("number", 1) - 1  # Parcelas começam do 1
            if index < 10:
                parcel_values[index] = installment.get("value")

                # Determinar a cor da célula com base no status e vencimento
                status = installment.get("status", "").upper()
                due_date_str = installment.get("due_date", "").split("T")[0]

                if due_date_str:
                    due_date = datetime.strptime(due_date_str, "%Y-%m-%d").date()
                    is_late = due_date < today and status == "PENDING"
                else:
                    is_late = False

                # Define a cor da célula
                if status == "ACQUITTED":
                    status_colors[index] = (0.6, 0.8, 0.6)  # Verde
                    font_colors[index] = (0, 0, 0)  # Preto
                elif is_late:
                    status_colors[index] = (1, 0.4, 0.4)  # Vermelho
                    font_colors[index] = (0, 0, 0)  # Preto
                else:
                    status_colors[index] = (1, 1, 0.6)  # Amarelo
                    font_colors[index] = (0, 0, 0)  # Preto

        # Adicionar os dados à lista
        filtered_installments.append([
            remove_acentos(seller.get("name")),
            remove_acentos(customer.get("name")),
            sale.get("number"),
            data,
            sale.get("total"),
            payment.get("method"),
            *parcel_values  # Adiciona as parcelas
        ])

        cell_formats.append({
            "colors": status_colors,
            "fonts": font_colors
        })  # Guarda as cores das células e fontes

# Autenticação com o Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file("token/api-planilha-rodrigo.json", scopes=scope)
client = gspread.authorize(creds)

# Abrir a planilha e aba específica
SHEET_NAME = "Comissão Time de Vendas TESTE"
WORKSHEET_NAME = "Carlos Louback"
sheet = client.open(SHEET_NAME).worksheet(WORKSHEET_NAME)

# Determinar a linha inicial para inserção
ultima_linha = len(sheet.get_all_values()) + 1
if ultima_linha < 9:
    ultima_linha = 9

# Definir o intervalo correto de inserção
num_linhas = len(filtered_installments)
range_inicio = ultima_linha
range_fim = range_inicio + num_linhas - 1
intervalo = f"A{range_inicio}:P{range_fim}"  # P é a última coluna das parcelas

# Inserir os valores na planilha
sheet.update(intervalo, filtered_installments)

# Aplicar cores nas células e nas fontes
requests = []
for i, format_data in enumerate(cell_formats):
    row = range_inicio + i  # Linha atual na planilha
    for j, (color, font_color) in enumerate(zip(format_data["colors"], format_data["fonts"])):
        if color:  # Se tiver cor definida
            # Garantir que font_color não seja None
            if font_color is None:
                font_color = (0, 0, 0)  # Preto como padrão se for None

            requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": sheet.id,
                        "startRowIndex": row - 1,
                        "endRowIndex": row,
                        "startColumnIndex": 6 + j,  # Começa na coluna G (índice 6)
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
                    "fields": "userEnteredFormat.backgroundColor,textFormat.foregroundColor"
                }
            })

# Enviar as requisições para colorir as células e ajustar a cor da fonte
if requests:
    body = {"requests": requests}
    sheet.spreadsheet.batch_update(body)

print("Dados inseridos, cores e fontes aplicadas com sucesso!")
