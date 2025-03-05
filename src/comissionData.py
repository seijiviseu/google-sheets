import gspread
from google.oauth2.service_account import Credentials
import pandas as pd


def setup_google_sheets(sheet_name, worksheet_name):
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("src/common/config/api-planilha-rodrigo.json", scopes=scope)
    client = gspread.authorize(creds)
    return client.open(sheet_name).worksheet(worksheet_name)

def load_data(sheet, csv_file):
    
    existing_data = sheet.get_all_values()[7:]
    if not existing_data:
        print("A planilha está vazia.")
        return None, None

    df_existing = pd.DataFrame(existing_data[1:], columns=existing_data[0])
    df_existing["Venda"] = df_existing["Venda"].astype(int)
    
    
    df_vendas = pd.read_csv(csv_file)
    df_vendas["venda"] = df_vendas["venda"].astype(int)
    
    return df_existing, df_vendas

def process_updates(sheet, df_existing, df_vendas):
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
                                "red": 66/255,
                                "green": 133/255,
                                "blue": 244/255
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
    
    return requests

def update_sheet(sheet, requests):
    if requests:
        sheet.spreadsheet.batch_update({"requests": requests})
        print("Células atualizadas com sucesso!")
    else:
        print("Nenhuma célula para atualizar.")


def main_comission(sheet_name="Comissão Time de Vendas TESTE", 
                  worksheet_name="Carlos Louback",
                  csv_file="src/common/input_files/comission_data.csv"):
    
    sheet = setup_google_sheets(sheet_name, worksheet_name)
    
    df_existing, df_vendas = load_data(sheet, csv_file)
    if df_existing is None or df_vendas is None:
        return
    
    requests = process_updates(sheet, df_existing, df_vendas)
    
    update_sheet(sheet, requests)

if __name__ == "__main__":
    main_comission()
