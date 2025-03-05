import sys
import os
sys.path.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src'))

from GoogleSheets_Automation import main_insertData
from comissionData import main_comission
from deleteData import main_delete

def main():
    sheet_name = "Comissão Time de Vendas TESTE"
    worksheet_name = "Carlos Louback"
    json_file = "src/common/input_files/data.json"
    csv_file = "src/common/input_files/comission_data.csv"
    delete_csv = "src/common/input_files/delete_data.csv"

    print("Verificando dados para deletar...")
    main_delete(sheet_name, worksheet_name, delete_csv)
    
    print("\nInserindo novos dados...")
    main_insertData()
    
    print("\nAtualizando dados de comissão...")
    main_comission(sheet_name, worksheet_name, csv_file)

if __name__ == "__main__":
    main()
