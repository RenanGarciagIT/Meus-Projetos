import pandas as pd
from datetime import datetime, timedelta
import os
import shutil
import pyautogui
import pyperclip
import time

engeman = pyperclip.copy("C:/Users/renan.araujo/OneDrive - FORNECEDORA/Área de Trabalho/Engeman - Atalho.lnk")

# Obter a data atual
data_atual = datetime.now() - timedelta(days=1)
data_atual = datetime(data_atual.year,data_atual.month, data_atual.day)
data_formatada_atual = str(data_atual.strftime("%d%m%Y"))
ano_atual = str(data_atual.strftime("%Y"))
mes_atual = str(data_atual.strftime("%m"))

# Definir a data para o primeiro dia do mês atual

data_inicial = datetime(data_atual.year, data_atual.month, 1)
data_formatada_inicial = str(data_inicial.strftime("%d%m%Y"))

# ACESSAR ENGEMAN

time.sleep(5)
pyautogui.hotkey('win','r')
time.sleep(5)
pyautogui.hotkey('ctrl', 'v')
time.sleep(5)
pyautogui.hotkey('Enter')
time.sleep(10)

# APERTAR EM ENTRAR
pyautogui.hotkey('enter')
time.sleep(15)

# IR PARA PROCESSO

pyautogui.moveTo(224,42)
time.sleep(5)
pyautogui.click(224,42)
time.sleep(5)

# IR PARA EDITOR DE RELATÓRIO

pyautogui.moveTo(249,67)
time.sleep(5)
pyautogui.click(249,67)
time.sleep(10)

# IR NO BINOCULO

pyautogui.moveTo(647,103)
time.sleep(5)
pyautogui.click(647,103)
time.sleep(5)

# IR PARA CÓDIGO , DIGITAR 12188 , ENTER , CLICAR NO CÓDIGO E CLICAR PELA SEGUNDA VEZ

pyautogui.write("12188")
time.sleep(5)
pyautogui.press('enter')
time.sleep(5)
pyautogui.moveTo(373,229)
time.sleep(5)
pyautogui.click(373,229)
time.sleep(5)
pyautogui.doubleClick(373,229)
time.sleep(5)

# INICIAR O RELATÓRIO

pyautogui.moveTo(250,316)
time.sleep(5)
pyautogui.doubleClick(250,316)
time.sleep(5)

# DATA INICIAL E DATA FINAL SERÁ NOs  2 DIAS ANTERIOR DE HOJE

caminho = pyperclip.copy('C:/Users/renan.araujo/OneDrive - FORNECEDORA/Área de Trabalho/APROPRIAÇÃO MÃO DE OBRA/BASE DE DADOS MENSALMENTE')

time.sleep(5)
pyautogui.moveTo(495,591)
time.sleep(5)
pyautogui.click(495,591)
time.sleep(5)
pyautogui.moveTo(402,609)
time.sleep(5)
pyautogui.click(402,609)
time.sleep(5)
pyautogui.moveTo(521,595)
time.sleep(5)
pyautogui.click(521,595)
time.sleep(5)
pyautogui.click(552,156)
time.sleep(5)
pyautogui.write(data_formatada_inicial)
time.sleep(5)
pyautogui.press('tab')
time.sleep(5)
pyautogui.write(data_formatada_atual)
time.sleep(5)
pyautogui.press('tab')
time.sleep(5)
pyautogui.moveTo(889,594)
time.sleep(5)
pyautogui.click(889,594)
time.sleep(120)
pyautogui.moveTo(167,34)
time.sleep(5)
pyautogui.click(167,34)
time.sleep(5)
pyautogui.moveTo(224,310)
time.sleep(5)
pyautogui.click(224,310)
time.sleep(5)
pyautogui.moveTo(712,528)
time.sleep(5)
pyautogui.click(712,528)
time.sleep(5)
pyautogui.hotkey('ctrl', 'l')
time.sleep(5)
pyautogui.hotkey('ctrl', 'v')
time.sleep(5)
pyautogui.press('enter')
time.sleep(5)
pyautogui.press('tab')
time.sleep(5)
pyautogui.press('tab')
time.sleep(5)
pyautogui.press('tab')
time.sleep(5)
pyautogui.press('tab')
time.sleep(5)
pyautogui.press('tab')
time.sleep(5)
pyautogui.press('tab')
time.sleep(5)
pyautogui.write("CSV" + " " + ano_atual + "_" + mes_atual)
time.sleep(5)
pyautogui.press('enter')
time.sleep(5)
pyautogui.press('tab')
time.sleep(5)
pyautogui.press('enter')
time.sleep(5)
pyautogui.hotkey('alt', 'f4')
time.sleep(5)
pyautogui.hotkey('alt', 'f4')
time.sleep(5)
pyautogui.hotkey('alt', 'f4')
time.sleep(5)
pyautogui.hotkey('enter')
time.sleep(5)
pyautogui.hotkey('enter')
time.sleep(5)
pyautogui.hotkey('enter')
time.sleep(5)

# LEITURA DOS DADOS

base_dados =pd.read_csv(f"C:/Users/renan.araujo/OneDrive - FORNECEDORA/Área de Trabalho/APROPRIAÇÃO MÃO DE OBRA/BASE DE DADOS MENSALMENTE/CSV {ano_atual}_{mes_atual}.csv",encoding='latin1', delimiter=';')

# RENOMENANDO AS COLUNAS
base_dados.rename(columns={'dadosFuncionario': 'FUNCIONÁRIO'}, inplace=True)
base_dados.rename(columns={'dadosHorasEscala': 'HORAS DA ESCALA'}, inplace=True)
base_dados.rename(columns={'dadosHorasOSS': 'HORAS DA OS'}, inplace=True)
base_dados.rename(columns={'dadosCustoPrevisto': 'CUSTO PREVISTO'}, inplace=True)
base_dados.rename(columns={'dadosCustoReal': 'CUSTO REAL'}, inplace=True)
base_dados.rename(columns={'dadosEficiencia': 'EFICIENCIA'}, inplace=True)

# RETIRANDO OS VALORES VAZIOS e as colunas descenessárias
base_dados = base_dados.dropna()
base_dados = base_dados.drop("CUSTO PREVISTO", axis=1)
base_dados = base_dados.drop("CUSTO REAL", axis=1)

# TRANSFORMANDO OS DADOS DE HORA PARA MINUTOS
def hora_para_minuto(x):
    horas, minutos = map(int , x.split(':'))
    resultado = horas + (minutos/60)
    resultador_formatado = round(resultado,2)
    resultado_virgula = str(resultador_formatado).replace(".",",")
    return resultado_virgula

base_dados['HORAS DA ESCALA'] = base_dados['HORAS DA ESCALA'].apply(hora_para_minuto)
base_dados['HORAS DA OS'] = base_dados['HORAS DA OS'].apply(hora_para_minuto)

# TRANSFORMANDO A COLUNA EFICIENCIA EM NÚMERO
def porcentagem(x):
    return str(x.strip("%"))

base_dados['EFICIENCIA'] = base_dados['EFICIENCIA'].apply(porcentagem)

# TRANSFORMANDO NUMERO EM MÊS
def numero_para_mes(x):
    meses = {
    "01": "janeiro",
    "02": "fevereiro",
    "03": "março",
    "04": "abril",
    "05": "maio",
    "06": "junho",
    "07": "julho",
    "08": "agosto",
    "09": "setembro",
    "10": "outubro",
    "11": "novembro",
    "12": "dezembro"}
    
    return meses.get(x)


# ADICIONANDO COLUNA DE MÊS E ANO
base_dados['ANO'] = int(ano_atual)
base_dados['MÊS'] = numero_para_mes(mes_atual)

# NOME DO ARQUIVO
nome = f"CSV {ano_atual}_{mes_atual}.csv"

# BAIXANDO O ARQUIVO
base_dados.to_csv(nome)

# PARTE QUE MOVE O ARQUIVO

caminho_original = f'C:/Users/renan.araujo/OneDrive - FORNECEDORA/Área de Trabalho/APROPRIAÇÃO MÃO DE OBRA/CSV {ano_atual}_{mes_atual}.csv'

pasta_destino = 'C:/Users/renan.araujo/OneDrive - FORNECEDORA/Área de Trabalho/APROPRIAÇÃO MÃO DE OBRA/BASE DE DADOS MENSALMENTE/'

if os.path.exists(os.path.join(pasta_destino, nome)):
    # Se existir, remove o arquivo de destino
     os.remove(os.path.join(pasta_destino, nome))

# Move o arquivo para a pasta de destino
shutil.move(caminho_original, pasta_destino)


# Pasta onde estão os arquivos CSV e variavel finais

pasta_destino = 'C:/Users/renan.araujo/OneDrive - FORNECEDORA/Área de Trabalho/APROPRIAÇÃO MÃO DE OBRA/BASE DE DADOS MENSALMENTE/'
pasta_destino_final = 'C:/Users/renan.araujo/OneDrive - FORNECEDORA/Área de Trabalho/APROPRIAÇÃO MÃO DE OBRA/BASE DE DADOS GERAL/'
nome_arquivo_final = "apropriação_mão_de_obra.csv"
caminho_original_final = f'C:/Users/renan.araujo/OneDrive - FORNECEDORA/Área de Trabalho/APROPRIAÇÃO MÃO DE OBRA/apropriação_mão_de_obra.csv'

# Lista para armazenar os DataFrames de cada arquivo
lista_base_dados = []

# Itera sobre os arquivos na pasta

for arquivo in os.listdir(pasta_destino):
    if arquivo.endswith('.csv'):
        arquivo_geral = os.path.join(pasta_destino, arquivo)
        base_dados_geral = pd.read_csv(arquivo_geral)
        lista_base_dados.append(base_dados_geral)

# Concatena os DataFrames em um único DataFrame e removendo colunas
base_dados_geral_final = pd.concat(lista_base_dados, ignore_index=False)
base_dados_geral_final = base_dados_geral_final.drop('BASE', axis=1)
base_dados_geral_final = base_dados_geral_final.drop('ATIVO', axis=1)

# Salva o DataFrame concatenado em um novo arquivo CSV

base_dados_geral_final.to_csv("apropriação_mão_de_obra.csv", index= False)

if os.path.exists(os.path.join(pasta_destino_final, nome_arquivo_final)):
    # Se existir, remove o arquivo de destino
     os.remove(os.path.join(pasta_destino_final, nome_arquivo_final))

# Move o arquivo para a pasta de destino
shutil.move(caminho_original_final, pasta_destino_final)
