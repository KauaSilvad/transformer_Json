import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import os

# Função para converter Excel para JSON
def excel_para_json(arquivo_excel):
    if not os.path.exists(arquivo_excel):
        print(f"Erro: O arquivo '{arquivo_excel}' não foi encontrado.")
        return None
    
    try:
        df = pd.read_excel(arquivo_excel)
        json_data = df.to_json(orient='records', lines=True)
        return json_data
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return None

# Função para autenticar e obter acesso ao Google Sheets
def autenticar_google_sheets():
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name('credenciais.json', scope)
        client = gspread.authorize(creds)
        return client
    except FileNotFoundError:
        print("Erro: O arquivo de credenciais 'credenciais.json' não foi encontrado.")
        return None
    except Exception as e:
        print(f"Erro ao autenticar com o Google Sheets: {e}")
        return None

# Função para converter dados do Google Sheets para JSON
def google_sheets_para_json(nome_arquivo_sheets, nome_planilha):
    client = autenticar_google_sheets()
    
    if client is None:
        return None
    
    try:
        sheet = client.open(nome_arquivo_sheets).worksheet(nome_planilha)
        dados = sheet.get_all_records()
        json_data = json.dumps(dados, indent=4)
        return json_data
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"Erro: O arquivo '{nome_arquivo_sheets}' não foi encontrado no Google Sheets.")
        return None
    except gspread.exceptions.WorksheetNotFound:
        print(f"Erro: A aba '{nome_planilha}' não foi encontrada.")
        return None
    except Exception as e:
        print(f"Erro ao acessar os dados do Google Sheets: {e}")
        return None

# Função principal para interagir com o usuário
def main():
    tipo_entrada = input("Você quer converter um arquivo Excel ou Google Sheets? (excel/sheets): ").lower()

    if tipo_entrada == "excel":
        arquivo_excel = input("Digite o caminho do arquivo Excel: ")
        json_resultado = excel_para_json(arquivo_excel)
        if json_resultado:
            print("Resultado em JSON:")
            print(json_resultado)
        else:
            print("Falha ao converter o arquivo Excel.")
    
    elif tipo_entrada == "sheets":
        arquivo_sheets = input("Digite o nome do arquivo do Google Sheets: ")
        planilha = input("Digite o nome da aba da planilha: ")
        json_resultado = google_sheets_para_json(arquivo_sheets, planilha)
        if json_resultado:
            print("Resultado em JSON:")
            print(json_resultado)
        else:
            print("Falha ao converter os dados do Google Sheets.")
    
    else:
        print("Opção inválida!")

# Executando o programa
if __name__ == "__main__":
    main()
