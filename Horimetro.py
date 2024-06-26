import pandas as pd
import pyodbc


# Configurações de conexão
server = 'xxxxx' 
database = 'sxxxxxx' 
username = 'xxxxxxxxxx' 
password = r"xxxxxxx"

conn_string = f'DRIVER=ODBC Driver 17 for SQL Server;SERVER={server};DATABASE={database};UID={username};PWD={password}'

# Pegando base de dados de máquinas paradas
try:
    conn = pyodbc.connect(conn_string)
    cursor = conn.cursor()

    # Exemplo: Executando uma query de seleção
    cursor.execute('''SELECT 
 FORMAT(EGM.COLACU.DATHOR,'dd/MM/yyyy') AS DATA,
 EGM.COLACU.VALOR AS HORÍMETRO,
 EGM.APLIC.TAG AS PRÉFIXO

 FROM EGM.COLACU

 -- REALIZANDO INNER JOIN PARA PEGAR APENAS O HORÍMETRO
 INNER JOIN EGM.PONCONACU
 ON EGM.COLACU.CODPONACU = EGM.PONCONACU.CODPONACU


 -- REALIZANDO INNER JOIN PARA PEGAR O PREFÍXO DAS MÁQUINAS
 INNER JOIN EGM.APLIC
 ON EGM.COLACU.CODAPL = EGM.APLIC.CODAPL

 -- APLICANDO O FILTRO
 WHERE 
 EGM.PONCONACU.CODUNI = 9 AND EGM.APLIC.ATIVO = 'S'

 ORDER BY
	EGM.APLIC.TAG DESC,
    EGM.COLACU.DATHOR ASC;''')
    
    # Extrair os resultados em uma lista de tuplas
    dados = cursor.fetchall()

     # Montando a base de dados com dicionário
    dicionario = []

    item_dicionario = None

     # Criando um loop para alimentar dicionário

    for linha in dados:
        item_dicionario = {
            'DATA': linha[0],
            'HORIMETRO_ACUMULADO': linha[1],
            'PREFIXO': linha[2]
        }
        dicionario.append(item_dicionario)
    
finally:
# Fechar a conexão
    if 'conn' in locals():
        conn.close()


# Transformando em um DataFrame
base_dados = pd.DataFrame(dicionario)

# Retirando a duplicada da data e prefixo
base_dados = base_dados.drop_duplicates(subset=['DATA', 'PREFIXO'])

# Criando as variaveis para a tratativa dos dados
HORIMETRO_PECORRIDO = []
PREFIXO_ANTERIOR = None
HORIMETRO_ANTERIOR = 0

# Pré-processamento dos dados

for index, row in base_dados.iterrows():

    if row['PREFIXO'] == PREFIXO_ANTERIOR or PREFIXO_ANTERIOR == None:
        HORIMETRO_PECORRIDO.append(row['HORIMETRO_ACUMULADO'] - HORIMETRO_ANTERIOR)
        HORIMETRO_ANTERIOR = row['HORIMETRO_ACUMULADO']
        PREFIXO_ANTERIOR = row['PREFIXO']
    
    elif row['PREFIXO'] != PREFIXO_ANTERIOR:
        HORIMETRO_PECORRIDO.append(0)
        HORIMETRO_ANTERIOR = row['HORIMETRO_ACUMULADO']
        PREFIXO_ANTERIOR = row['PREFIXO']


base_dados['HORIMETRO_PECORRIDO'] = HORIMETRO_PECORRIDO

# Subistituindo "." para "," para power bi reconhecer que é um número float.

base_dados['HORIMETRO_PECORRIDO'] = base_dados['HORIMETRO_PECORRIDO'].astype(str)
base_dados['HORIMETRO_PECORRIDO'] = base_dados['HORIMETRO_PECORRIDO'].str.replace('.', ',')

base_dados['HORIMETRO_ACUMULADO'] = base_dados['HORIMETRO_ACUMULADO'].astype(str)
base_dados['HORIMETRO_ACUMULADO'] = base_dados['HORIMETRO_ACUMULADO'].str.replace('.', ',')

print(base_dados.head(30))
