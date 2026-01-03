import pyodbc
import os
import pandas as pd
from datetime import datetime, timedelta
import warnings
import configparser

config = configparser.ConfigParser()
config.read('./config/config.ini', encoding='utf-8')

# Silencia o aviso específico do pandas sobre a conexão DBAPI2
warnings.filterwarnings("ignore", category=UserWarning, module='pandas.io.sql')

# Configurações de Conexão
CONN_STR = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    f"SERVER={config['database']['host']},{config['database']['port']}\SQLEXPRESS;"
    f"DATABASE={config['database']['database']};"
    f"UID={config['database']['user']};"
    f"PWD={config['database']['password']};"
)

# query para buscar leads de uma data específica
FIND_LEADS_QUERY = config['query']['find_leads']

# Função para exibir o menu e obter a opção do usuário
def get_menu_option():
    """Exibe o menu e retorna a opção escolhida pelo usuário."""
    print("="*40)
    print("       Extração de Leads Diária        ")
    print("="*40)
    print("1 - Extrair leads de ONTEM")
    print("2 - Extrair leads de uma DATA ESPECÍFICA")
    print("3 - Sair")
    return input("Opção (1, 2 ou 3): ")

# Função para perguntar se o usuário deseja realizar uma nova consulta
def new_query():
    """Pergunta ao usuário se deseja realizar uma nova consulta."""
    print("Deseja realizar uma nova consulta?")
    print("1 - Sim")
    print("2 - Não")
    new_query_option = input("Opção (1 ou 2): ")
    if new_query_option == "1":
            clear_screen()
            run_query()
    elif new_query_option == "2":
        print("Encerrando o programa. Até mais!")
    else:
        print("Opção inválida. Tente Novamente.")
        new_query()

# Função principal para executar a extração
def execute_extraction(target_date):
    """Executa a query e salva em Excel com base na data fornecida."""
    try:
        print(f"Conectando ao banco e buscando dados de: {target_date}...")
        
        with pyodbc.connect(CONN_STR) as conn:
            # O pandas aceita parâmetros para a query para evitar SQL Injection
            df = pd.read_sql(FIND_LEADS_QUERY, conn, params=[target_date])

        if df.empty:
            print(f"Atenção: Nenhum dado encontrado para a data {target_date}.")
            return

        file_name = f"leads_{target_date}.xlsx"
        df.to_excel(file_name, index=False, engine="openpyxl")
        print(f"Sucesso! Arquivo gerado: {file_name}")
        print(f"Total de registros: {len(df)}")

    except Exception as e:
        print(f"Erro durante a execução: {e}")

# Função para limpar a tela
def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

# Função para gerenciar o fluxo do programa
def run_query():
    """Gerencia o fluxo do programa."""
    option = get_menu_option()
    if option == "1":
        target_date = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
        execute_extraction(target_date)
        new_query()

    elif option == "2":
        date_input = input("Digite a data (AAAA-MM-DD): ")
        try:
            # Valida se o formato está correto
            valid_date = datetime.strptime(date_input, "%Y-%m-%d").date()
            execute_extraction(str(valid_date))
            new_query()
        except ValueError:
            print("Formato de data inválido! Use AAAA-MM-DD (Ex: 2023-10-25).")
            run_query()

    elif option == "3":
        print("Encerrando o programa. Até mais!")

    else:
        print("Opção inválida. Tente Novamente.")
        run_query()

# Início do Programa
run_query()