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
    #print(insert_query)
    try:
        for _, row in df.iterrows():
            cursor.execute(insert_query, tuple(row))
        connection.commit()
        print(f"Dados inseridos na tabela {tabela_banco} com sucesso.")
    except pyodbc.Error as e:
        print(f"Erro ao inserir dados na tabela {tabela_banco}: {e}")


def remove_duplicatas(connection):
    print('REMOVENDO DATAS DUPLICADAS DAS TABELAS ABAIXO')
    tabelas = {
        "SR_VisaoGeralDominio": "Data",
        "SR_VisitasSite": "Mes",
        "SR_TaxaRejeicao": "Mes",
        "SR_MediaDuracaoVisita": "Mes",
    }
    
    for tabela, coluna_chave in tabelas.items():
        print(f"Removendo duplicatas da tabela {tabela}...")
        try:
            cursor = connection.cursor()

            # SQL para remover duplicatas mantendo apenas a linha com a Data_Extracao mais recente
            sql_remover_duplicatas = f"""
            WITH CTE AS (
                SELECT *,
                    ROW_NUMBER() OVER (PARTITION BY {coluna_chave} ORDER BY Data_Extracao DESC) AS rn
                FROM {tabela}
            )
            DELETE FROM CTE
            WHERE rn > 1;
            """
            
            cursor.execute(sql_remover_duplicatas)
            connection.commit()
            print(f"Duplicatas removidas da tabela {tabela}.")
        except pyodbc.Error as e:
            print(f"Erro ao remover duplicatas da tabela {tabela}: {e}")
        finally:
            cursor.close()


def job():
    driver = "ODBC Driver 17 for SQL Server"
    server = "00"  # Substitua pelo seu servidor
    user = "00"  # Substitua pelo seu usuário
    password = "00"  # Substitua pela sua senha
    database = "00"
    port = 00

    connection = create_connection(driver, server, database, user, password, port)

    if connection:
        caminho_excel = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ExcelTratado", "DadosTratados.xlsx")

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
                    "Categoria": "Categoria",
                    "Data_Extracao": "Data_Extracao"
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
                    "Soma_Total": "Soma_Total",
                    "Data_Extracao": "Data_Extracao"
                    
                },
                "VisitasSite": {
                    "Mes": "Mes",
                    "bagaggio": "bagaggio",
                    "lepostiche": "lepostiche",
                    "inovathi": "inovathi",
                    "sestini": "sestini",
                    "gocase": "gocase",
                    "Data_Extracao": "Data_Extracao"
                },
                "TaxaRejeicao": {
                    "Mes": "Mes",
                    "bagaggio": "bagaggio",
                    "lepostiche": "lepostiche",
                    "inovathi": "inovathi",
                    "sestini": "sestini",
                    "gocase": "gocase",
                    "Data_Extracao": "Data_Extracao"
                },
                "MediaDuracaoVisita": {
                    "Mes": "Mes",
                    "bagaggio": "bagaggio",
                    "lepostiche": "lepostiche",
                    "inovathi": "inovathi",
                    "sestini": "sestini",
                    "gocase": "gocase",
                    "Data_Extracao": "Data_Extracao"
                },
                "JornadaTrafego": {
                    "Canal": "Canal",
                    "bagaggio.com.br": "bagaggio.com.br",
                    "lepostiche.com.br": "lepostiche.com.br",
                    "inovathi.com.br": "inovathi.com.br",
                    "sestini.com.br": "sestini.com.br",
                    "gocase.com.br": "gocase.com.br",
                    "Data_Extracao": "Data_Extracao"
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
                    "Categoria": "Categoria",
                    "Data_Extracao": "Data_Extracao"
                },
                "LacunasBacklinks": {
                    "Domain": "Domain",
                    "Domain ascore": "Domain ascore",
                    "bagaggio": "bagaggio",
                    "lepostiche": "lepostiche",
                    "inovathi": "inovathi",
                    "sestini": "sestini",
                    "gocase": "gocase",
                    "Matches": "Matches",
                    "Data_Extracao": "Data_Extracao"
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
                    
                insert_data_from_df(connection, df, tabela)
        else:
            print(f"Arquivo {caminho_excel} não encontrado.")
    else:
        print("Código está correto, mas não foi possível estabelecer a conexão.")
    
    remove_duplicatas(connection)
    print("Fim da execução.")

if __name__ == "__main__":
    job()