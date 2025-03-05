import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
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
    if "Venda" in df_existing.columns:
        df_existing["Venda"] = pd.to_numeric(df_existing["Venda"], errors="coerce")

    df_delete = pd.read_csv(csv_file)
    if "Venda" in df_delete.columns:
        df_delete["Venda"] = pd.to_numeric(df_delete["Venda"], errors="coerce")

    return df_existing, df_delete


def delete_rows(sheet, df_existing, df_delete):
    deleted_count = 0
    for venda in df_delete["Venda"]:
        index_list = df_existing[df_existing["Venda"] == venda].index.tolist()
        
        for idx in sorted(index_list, reverse=True):
            row_to_delete = idx + 9
            sheet.delete_rows(row_to_delete)
            deleted_count += 1
            print(f"Deletada a venda {venda} na linha {row_to_delete}")
    
    return deleted_count


def main_delete(sheet_name="Comissão Time de Vendas TESTE",
                worksheet_name="Carlos Louback",
                csv_file="src/common/input_files/delete_data.csv"):
    
    sheet = setup_google_sheets(sheet_name, worksheet_name)
    
    df_existing, df_delete = load_data(sheet, csv_file)
    if df_existing is None or df_delete is None:
        return
    
    deleted_count = delete_rows(sheet, df_existing, df_delete)
    
    if deleted_count > 0:
        print(f"\nTotal de {deleted_count} linha(s) deletada(s) com sucesso!")
    else:
        print("\nNenhuma linha encontrada para deletar.")


if __name__ == "__main__":
    main_delete()
