import gspread
from google.oauth2.service_account import Credentials
import pandas as pd


scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file("src/token/api-planilha-rodrigo.json", scopes=scope)
client = gspread.authorize(creds)

SHEET_NAME = "Comissão Time de Vendas TESTE"
WORKSHEET_NAME = "Carlos Louback"
sheet = client.open(SHEET_NAME).worksheet(WORKSHEET_NAME)

csv_file = "src/data/comission_data.csv"
df_vendas = pd.read_csv(csv_file)

existing_data = sheet.get_all_values()[7:]
if not existing_data:
    print("A planilha está vazia.")
    exit()

df_existing = pd.DataFrame(existing_data[1:], columns=existing_data[0])
df_existing["Venda"] = df_existing["Venda"].astype(int)
df_vendas["venda"] = df_vendas["venda"].astype(int)

requests = []

for _, row_csv in df_vendas.iterrows():
    venda_id = int(row_csv["venda"])  
    parcela_quitada = int(row_csv["parcela"]) - 1  

    venda_index = df_existing.index[df_existing["Venda"] == venda_id].tolist()
    
    if venda_index:
        index = venda_index[0]
        row_index = index + 9  

        coluna_parcela = 6 + parcela_quitada

        requests.append({
            "repeatCell": {
                "range": {
                    "sheetId": sheet.id,
                    "startRowIndex": row_index - 1,
                    "endRowIndex": row_index,
                    "startColumnIndex": coluna_parcela,
                    "endColumnIndex": coluna_parcela + 1
                },
                "cell": {
                    "userEnteredFormat": {
                        "backgroundColor": {
                            "red": 0.3,
                            "green": 0.5,
                            "blue": 1.0
                        },
                        "textFormat": {
                            "foregroundColor": {
                                "red": 1.0,
                                "green": 1.0,
                                "blue": 1.0
                            }
                        }
                    }
                },
                "fields": "userEnteredFormat(backgroundColor,textFormat.foregroundColor)"
            }
        })

if requests:
    sheet.spreadsheet.batch_update({"requests": requests})
    print("Células atualizadas com sucesso!")
else:
    print("Nenhuma célula para atualizar.")
