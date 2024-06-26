import pyodbc
import pandas as pd
from datetime import datetime as dt
import os
import shutil
import time

# Configurações de conexão
server = 'XXXXXXX' 
database = 'XXXXXXX 
username = 'XXXXXXXX' 
password = r"XXXXXXXXX"


conn_string = f'DRIVER=ODBC Driver 17 for SQL Server;SERVER={server};DATABASE={database};UID={username};PWD={password}'

# Pegando base de dados de máquinas paradas
try:
    conn = pyodbc.connect(conn_string)
    cursor = conn.cursor()

    # Exemplo: Executando uma query de seleção
    cursor.execute('SELECT * FROM EGM.T_DISPONIBILIDADE_EQUIPAMENTO')

    # Extrair os resultados em uma lista de tuplas
    dados = cursor.fetchall()

    # Montando a base de dados com dicionário
    dicionario = []

    item_dicionario = None

    # Criando um loop para alimentar dicionário

    for linha in dados:
        item_dicionario = {
            'PREFIXO': linha[0],
            'DATA': linha[1].strftime("%d-%m-%Y"),
            'PARADA':linha[2],
            'CENTRO_CUSTO': linha[3],
            'OPERACIONAL': linha[4],
            'STATUS': linha[5]
        }

        dicionario.append(item_dicionario)

finally:
    # Fechar a conexão
    if 'conn' in locals():
        conn.close()


# Pegando base de dados de máquinas paradas apenas avarias terceiras
try:
    conn = pyodbc.connect(conn_string)
    cursor = conn.cursor()

    # Exemplo: Executando uma query de seleção
    cursor.execute('SELECT * FROM EGM.T_DISPONIBILIDADE_EQUIPAMENTO_AVARIA_TERCEIRA')

    # Extrair os resultados em uma lista de tuplas
    dados_avarias = cursor.fetchall()

    # Montando a base de dados com dicionário
    dicionario_avarias = []

    item_dicionario_avarias = None

    # Criando um loop para alimentar dicionário

    for linha in dados_avarias:
        item_dicionario_avarias = {
            'PREFIXO': linha[0],
            'DATA': linha[1].strftime("%d-%m-%Y"),
            'PARADA':linha[2],
            'CENTRO_CUSTO': linha[3],
            'OPERACIONAL': linha[4],
            'STATUS': linha[5]
        }

        dicionario_avarias.append(item_dicionario_avarias)

finally:
    # Fechar a conexão
    if 'conn' in locals():
        conn.close()



# Transformando em um DataFrame
nova_base = pd.DataFrame(dicionario)
nova_base_avarias = pd.DataFrame(dicionario_avarias)

# Realizando merge das base de dados

resultado = pd.merge(nova_base , nova_base_avarias, on=['PREFIXO' , 'DATA'])
resultado['PARADA'] = resultado['PARADA_x'] - resultado['PARADA_y']

# Filtrando apenas paradas, retirando avarias terceria e tratando os dados
condição = lambda row: "PARADA" if row['PARADA'] > 0 else "FUNCIONANDO"
resultado['STATUS'] = resultado.apply(condição, axis=1)
resultado['DATA'] = pd.to_datetime(resultado['DATA'], format='%d-%m-%Y')
resultado = resultado.drop(['PARADA_x', 'PARADA_y', 'CENTRO_CUSTO_y','OPERACIONAL_y','STATUS_y','STATUS_x'  ], axis=1)
resultado.rename(columns={'CENTRO_CUSTO_x' : 'CENTRO_CUSTO', 'OPERACIONAL_x':'OPERACIONAL'}, inplace=True)
resultado = resultado[resultado['PARADA'] == 24]
resultado = resultado.sort_values(by=['PREFIXO', 'DATA'])

# Criando loop para datas consecutivas

data_anterior = None
dicionario_dias = []
prefixo_anterior = None
dias_consecutivos = 1
centro_custo_anterior = None

for index, row in resultado.iterrows():

    if row['PREFIXO'] != prefixo_anterior or row['CENTRO_CUSTO'] != centro_custo_anterior:
        dicionario_dias.append(1)
        data_anterior = row['DATA']
        prefixo_anterior = row['PREFIXO']
        centro_custo_anterior = row['CENTRO_CUSTO']
    
    elif row['PREFIXO'] == prefixo_anterior:
        
        if (row["DATA"] - data_anterior).days == 1:
            if dicionario_dias[-1] == 15:
                dicionario_dias.append(1)
            else:
                dicionario_dias.append(dicionario_dias[-1] + 1)
            data_anterior = row['DATA']

        else:
            dicionario_dias.append(1)
            data_anterior = row['DATA']

# Adicionando nova coluna
resultado['DIAS CONSECUTIVOS'] = dicionario_dias

# Tratando a coluna parada
resultado['PARADA'] = resultado['PARADA'].astype(str)
resultado['PARADA'] = resultado['PARADA'].str.replace('.', ',')


# Baixando o arquivo em CSV e nomeando
nome_arquivo_final = str("MÁQUINA PARADA +15 DIAS.csv")
nova_base.to_csv(nome_arquivo_final, index=False)

# Transportando o arquivo na pasta BASE DADOS 

caminho_original_final = f"C:/Users/renan.araujo/OneDrive - FORNECEDORA/Área de Trabalho/MÁQUINA PARADA +15 DIAS/MÁQUINA PARADA +15 DIAS.csv"
pasta_destino_final = 'C:/Users/renan.araujo/OneDrive - FORNECEDORA/Área de Trabalho/MÁQUINA PARADA +15 DIAS/BASE DADOS/'

if os.path.exists(os.path.join(pasta_destino_final, nome_arquivo_final)):
    # Se existir, remove o arquivo de destino
     os.remove(os.path.join(pasta_destino_final, nome_arquivo_final))

# Move o arquivo para a pasta de destino
shutil.move(caminho_original_final, pasta_destino_final)

time.sleep(5)

print("*"*30)
# Código exercutado
if os.path.exists(os.path.join(pasta_destino_final, nome_arquivo_final)):
    print("Exercutado com sucesso!!!")
else:
    print("Ocorreu um erro em transportar o arquivo")

print("*"*30)

print(resultado)
