import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd

# Escopo de acesso
scope = ['https://www.googleapis.com/auth/spreadsheets',
         'https://www.googleapis.com/auth/drive']

# Caminho para o arquivo JSON de credenciais
# Substitua 'path_to_your_credentials.json' pelo caminho para o seu arquivo JSON de credenciais
json_keyfile_path = 'path_to_your_credentials.json'

# Autenticando
credentials = ServiceAccountCredentials.from_json_keyfile_name(
    json_keyfile_path, scope)
client = gspread.authorize(credentials)

# Função para sincronizar Excel para Google Sheets


def sync_excel_to_google_sheets(excel_file_path, spreadsheet_id, sheet_name):
    try:
        # Abrir a planilha pelo ID
        spreadsheet = client.open_by_key(spreadsheet_id)
        sheet = spreadsheet.worksheet(sheet_name)

        # Ler o arquivo Excel local
        df = pd.read_excel(excel_file_path)

        df.fillna(value='', inplace=True)  # tratar o Valor NaN

        # Atualizar a planilha do Google Sheets
        sheet.clear()  # Limpar conteúdo existente
        sheet.update([df.columns.values.tolist()] + df.values.tolist())

        print("Sincronização concluída com sucesso.")
    except Exception as e:
        print(f"Ocorreu um erro durante a sincronização: {e}")


# ID da planilha no Google Sheets
spreadsheet_id = 'your_spreadsheet_id_here'

# Caminho local do arquivo Excel
excel_file_path = 'path_to_your_excel_file.xlsx'

# Nome da folha que você deseja sincronizar
sheet_name = 'your_sheet_name_here'

# Chamada para sincronizar
sync_excel_to_google_sheets(excel_file_path, spreadsheet_id, sheet_name)
