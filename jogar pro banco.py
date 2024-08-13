import os
import pandas as pd
import pyodbc

def create_connection(driver, server, database, user, password, port):
    try:
        connection = pyodbc.connect(
            f'DRIVER={{{driver}}};'
            f'SERVER={server},{port};'
            f'DATABASE={database};'
            f'UID={user};'
            f'PWD={password}'
        )
        print("Connection to SQL Server successful")
        return connection
    except pyodbc.Error as e:
        print(f"The error '{e}' occurred")
        return None

def truncate_table(connection, table_name):
    try:
        with connection.cursor() as cursor:
            cursor.execute(f"TRUNCATE TABLE {table_name}")
            connection.commit()
            print(f"Tabela {table_name} foi truncada com sucesso.")
    except pyodbc.Error as e:
        print(f"O erro foi: {e}")

def clean_and_convert_dataframe(df, column_mapping, expected_columns):
    df.rename(columns=column_mapping, inplace=True)

    for col, dtype in expected_columns.items():
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            if dtype == 'int':
                df[col] = df[col].astype(int)
            elif dtype == 'float':
                df[col] = df[col].round(2)  # Arredondar para 2 casas decimais

    return df

def insert_data_from_df(connection, df, tabela_banco):
    cursor = connection.cursor()
    columns = ', '.join([f'[{col}]' for col in df.columns])
    placeholders = ', '.join(['?' for _ in df.columns])
    insert_query = f"INSERT INTO {tabela_banco} ({columns}) VALUES ({placeholders})"
    print(insert_query)
    try:
        for _, row in df.iterrows():
            cursor.execute(insert_query, tuple(row))
        connection.commit()
        print(f"Dados inseridos na tabela {tabela_banco} com sucesso.")
    except pyodbc.Error as e:
        print(f"Erro ao inserir dados na tabela {tabela_banco}: {e}")

def job():
    driver = "ODBC Driver 17 for SQL Server"
    server = "000000"  # your server
    user = "000000"  # your username
    password = "000000"  # your password
    database = "000000" # your database
    port = 00000 # your port

    connection = create_connection(driver, server, database, user, password, port)

    if connection:
        caminho_excel = os.path.join(os.getcwd(), "PRIVATE-SEMRUSH-Automation", "ExcelTratado", "DadosTratados.xlsx")

        if os.path.exists(caminho_excel):
            column_mappings = {
                "VisaoGeralPalavrasChave": {
                    "Keyword": "Keyword",
                    "Intent": "Intent",
                    "Volume": "Volume",
                    "Trend": "Trend",
                    "Keyword Difficulty": "Keyword Difficulty",
                    "CPC (BRL)": "CPC (BRL)",
                    "SERP Features": "SERP Features",
                    "Categoria": "Categoria"
                },
                "VisaoGeralDominio": {
                    "Target": "Target",
                    "Mes": "Mes",
                    "Organic Keywords": "Organic Keywords",
                    "Organic Traffic": "Organic Traffic",
                    "Organic Traffic Cost": "Organic Traffic Cost",
                    "Paid Keywords": "Paid Keywords",
                    "Paid Traffic": "Paid Traffic",
                    "Paid Traffic Cost": "Paid Traffic Cost",
                    "Soma_Total": "Soma_Total"
                },
                "VisitasSite": {
                    "Mes": "Mes",
                    "bagaggio": "bagaggio",
                    "lepostiche": "lepostiche",
                    "inovathi": "inovathi",
                    "sestini": "sestini",
                    "gocase": "gocase"
                },
                "TaxaRejeicao": {
                    "Mes": "Mes",
                    "bagaggio": "bagaggio",
                    "lepostiche": "lepostiche",
                    "inovathi": "inovathi",
                    "sestini": "sestini",
                    "gocase": "gocase"
                },
                "MediaDuracaoVisita": {
                    "Mes": "Mes",
                    "bagaggio": "bagaggio",
                    "lepostiche": "lepostiche",
                    "inovathi": "inovathi",
                    "sestini": "sestini",
                    "gocase": "gocase"
                },
                "JornadaTrafego": {
                    "Canal": "Canal",
                    "bagaggio.com.br": "bagaggio.com.br",
                    "lepostiche.com.br": "lepostiche.com.br",
                    "inovathi.com.br": "inovathi.com.br",
                    "sestini.com.br": "sestini.com.br",
                    "gocase.com.br": "gocase.com.br"
                },
                "LacunasPalavrasChave": {
                    "Keyword": "Keyword",
                    "Search Volume": "Search Volume",
                    "Keyword Difficulty": "Keyword Difficulty",
                    "CPC": "CPC",
                    "Competition": "Competition",
                    "Results": "Results",
                    "Keyword Intents": "Keyword Intents",
                    "bagaggio (pages)": "bagaggio (pages)",
                    "lepostiche (pages)": "lepostiche (pages)",
                    "inovathi (pages)": "inovathi (pages)",
                    "sestini (pages)": "sestini (pages)",
                    "gocase (pages)": "gocase (pages)",
                    "Categoria": "Categoria"
                },
                "LacunasBacklinks": {
                    "Domain": "Domain",
                    "Domain ascore": "Domain ascore",
                    "bagaggio": "bagaggio",
                    "lepostiche": "lepostiche",
                    "inovathi": "inovathi",
                    "sestini": "sestini",
                    "gocase": "gocase",
                    "Matches": "Matches"
                }
            }

            abas_tabelas = {
                "VisaoGeralPalavrasChave": "SR_VisaoGeralPalavrasChave",
                "VisaoGeralDominio": "SR_VisaoGeralDominio",
                "VisitasSite": "SR_VisitasSite",
                "TaxaRejeicao": "SR_TaxaRejeicao",
                "MediaDuracaoVisita": "SR_MediaDuracaoVisita",
                "JornadaTrafego": "SR_JornadaTrafego",
                "LacunasPalavrasChave": "SR_LacunasPalavrasChave",
                "LacunasBacklinks": "SR_LacunasBacklinks"
            }

            

            for aba, tabela in abas_tabelas.items():
                df = pd.read_excel(caminho_excel, sheet_name=aba)

                if tabela in column_mappings:
                    df = clean_and_convert_dataframe(df, column_mappings[aba], {})

                truncate_table(connection, tabela)
                insert_data_from_df(connection, df, tabela)
        else:
            print(f"Arquivo {caminho_excel} não encontrado.")
    else:
        print("Código está correto, mas não foi possível estabelecer a conexão.")
    print("Fim da execução.")

if __name__ == "__main__":
    job()