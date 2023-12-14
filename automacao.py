from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import googleapiclient.discovery
import googleapiclient.errors
import pyautogui
from google.oauth2 import service_account
from datetime import datetime, timedelta
from time import sleep
import time
import pandas as pd
import os
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
import gspread_formatting
from gspread_formatting import Color
import gspread
import shutil
from io import StringIO
import re
import numpy as np

# EXPORTAR PLANILHA CARGAS sem_recebedor

# Abrir navegador

driver = webdriver.Chrome()
driver.maximize_window()
driver.get('URL do site aqui')

# Fazer login
campo_usuario = driver.find_element(By.XPATH, '//*[@id="user"]')
campo_usuario.send_keys('Usuário')
campo_senha = driver.find_element(By.XPATH, '//*[@id="password"]')
campo_senha.send_keys('Senha')
botao_entrar = driver.find_element(By.XPATH, '//*[@id="acessar"]')
botao_entrar.click()
time.sleep(6)

# Filtros para exportação
atalho = driver.find_element(By.XPATH, '//*[@id="codigo_opcao"]')
atalho.send_keys('126')
atalho.send_keys(Keys.ENTER)
time.sleep(1)

hoje = datetime.now()
semana = hoje.weekday()

if semana == 0:
     ontem = hoje - timedelta(days=3)
     data_final = hoje - timedelta(days=1)
else:
     ontem = hoje - timedelta(days=1)
     data_final = hoje - timedelta(days=1)

data = ontem.date()
data2 = data_final.date()
textodata = data.strftime("%d%m%Y")
textodata3 = data2.strftime("%d%m%Y")

data_in = driver.find_element(By.XPATH, '//*[@id="data_inicial"]')
data_fi = driver.find_element(By.XPATH, '//*[@id="data_final"]')

data_in.click()
data_in.send_keys(textodata)
time.sleep(0.5)
data_fi.click()
data_fi.send_keys(textodata3)
time.sleep(0.5)

dropdown_situacao = driver.find_element(By.XPATH, '//*[@id="situacao"]')
opcoes = Select(dropdown_situacao)
opcoes.select_by_visible_text('SEM RECEBEDOR')
time.sleep(0.5)

pesquisar = driver.find_element(By.XPATH, '//*[@id="PESQUISAR"]')
pesquisar.click()
time.sleep(10)

excel = driver.find_element(By.XPATH, '//*[@id="lista-minuta-excel"]')
excel.click()
time.sleep(8)


# Mover planilha

# Caminho para a pasta onde os arquivos Excel estão localizados
pasta_origem = 'Caminho da pasta de origem'

# Caminho para a pasta de destino
pasta_destino = 'Caminho da pasta de destino'

# Lista todos os arquivos na pasta de origem
arquivos = os.listdir(pasta_origem)

# Procura pelo arquivo Excel (.xlsx) na lista de arquivos e move para a pasta de destino
for arquivo in arquivos:
     if arquivo.lower().endswith('.xls'):
         caminho_arquivo = os.path.join(pasta_origem, arquivo)
         # Move o arquivo para a pasta de destino
         shutil.move(caminho_arquivo, pasta_destino)
         break  # Para de procurar após encontrar o primeiro arquivo Excel

time.sleep(5)

# Lista todos os arquivos na pasta de origem
n_arquivos = os.listdir(pasta_destino)

# Procura pelo arquivo Excel (.xlsx) na lista de arquivos
for n_arquivo in n_arquivos:
    if n_arquivo.lower().endswith('.xls'):
        caminho2_arquivo = os.path.join(pasta_destino, n_arquivo)

        # Novo nome para o arquivo Excel
        novo_nome = 'cargas_SR.xls'  # Defina o novo nome desejado
        novo_caminho_arquivo = os.path.join(pasta_destino, novo_nome)
        os.rename(caminho2_arquivo, novo_caminho_arquivo)

        break  # Para de procurar após encontrar o primeiro arquivo Excel

time.sleep(6)

# Lê o conteúdo do arquivo 'cargas_sem_recebedor.xls' para uma string
with open('cargas_SR.xls', 'r', encoding='utf-8') as file:
    html_content = file.read()

# Usa StringIO para criar um objeto de leitura da string
html_io = StringIO(html_content)

# Lê as tabelas do HTML usando o objeto StringIO
tables = pd.read_html(html_io)


# Agora você pode acessar as tabelas como DataFrames
for idx, table in enumerate(tables):
    # Salva a tabela como arquivo XLSX
    table.to_excel('cargas_sem_recebedor.xlsx', index=False)

time.sleep(3)

# EXPORTAR PLANILHA HOJE

# Filtros para exportação
atalho = driver.find_element(By.XPATH, '//*[@id="codigo_opcao"]')
atalho.send_keys('126')
atalho.send_keys(Keys.ENTER)
time.sleep(1)

hoje = datetime.now()
semana = hoje.weekday()

if semana == 0:
    ontem = hoje - timedelta(days=15)
    data_final = hoje
else:
    ontem = hoje - timedelta(days=16)
    data_final = hoje

data = ontem.date()
data2 = data_final.date()
textodata2 = data.strftime("%d%m%Y")
textodata4 = data2.strftime("%d%m%Y")

data_in = driver.find_element(By.XPATH, '//*[@id="data_inicial"]')
data_fi = driver.find_element(By.XPATH, '//*[@id="data_final"]')

data_in.click()
data_in.send_keys(textodata2)
time.sleep(0.5)
data_fi.click()
data_fi.send_keys(textodata4)
time.sleep(0.5)

dropdown_situacao = driver.find_element(By.XPATH, '//*[@id="situacao"]')
opcoes = Select(dropdown_situacao)
opcoes.select_by_visible_text('ENTREGUE')
time.sleep(0.5)

pesquisar = driver.find_element(By.XPATH, '//*[@id="PESQUISAR"]')
pesquisar.click()
time.sleep(10)

excel = driver.find_element(By.XPATH, '//*[@id="lista-minuta-excel"]')
excel.click()

time.sleep(8)

# Mover planilha

# Lista todos os arquivos na pasta de origem
arquivos = os.listdir(pasta_origem)

# Procura pelo arquivo Excel (.xlsx) na lista de arquivos e move para a pasta de destino
for arquivo in arquivos:
    if arquivo.lower().endswith('.xls'):
        caminho_arquivo = os.path.join(pasta_origem, arquivo)


#         # Move o arquivo para a pasta de destino
        shutil.move(caminho_arquivo, pasta_destino)

        break  # Para de procurar após encontrar o primeiro arquivo Excel

time.sleep(5)

# Lista todos os arquivos na pasta de origem
n_arquivos = os.listdir(pasta_destino)

# Contador para acompanhar o número de arquivos .xlsx encontrados
contador_xlsx = 0

# Procura pelo arquivo Excel (.xlsx) na lista de arquivos
for n_arquivo in n_arquivos:
    if n_arquivo.lower().endswith('.xls'):
        contador_xlsx += 1  # Incrementa o contador

#         # Verifica se é o segundo arquivo .xlsx encontrado
        if contador_xlsx == 2:
            caminho2_arquivo = os.path.join(pasta_destino, n_arquivo)

#              # Novo nome para o arquivo Excel
            novo_nome = 'cargas_E.xls'  # Defina o novo nome desejado
            novo_caminho_arquivo = os.path.join(pasta_destino, novo_nome)
            os.rename(caminho2_arquivo, novo_caminho_arquivo)

            break  # Para de procurar após encontrar o primeiro arquivo Excel


time.sleep(6)

# Lê o conteúdo do arquivo 'cargas_sem_recebedor.xls' para uma string
with open('cargas_E.xls', 'r', encoding='utf-8') as file:
    html_content = file.read()

# Usa StringIO para criar um objeto de leitura da string
html_io = StringIO(html_content)

# Lê as tabelas do HTML usando o objeto StringIO
tables = pd.read_html(html_io)

# Agora você pode acessar as tabelas como DataFrames
for idx, table in enumerate(tables):
    # Salva a tabela como arquivo XLSX
    table.to_excel('cargas_entregues.xlsx', index=False)

time.sleep(6)

# Extrair terceira planilha 

driver.get('URL site')

# Fazer Login
campo_usuario2 = driver.find_element(By.XPATH, '//*[@id="rcmloginuser"]')
campo_usuario2.send_keys('Login e-mail')
campo_senha2 = driver.find_element(By.XPATH, '//*[@id="rcmloginpwd"]')
campo_senha2.send_keys('Senha')
botao_entrar2 = driver.find_element(By.XPATH, '//*[@id="rcmloginsubmit"]')
botao_entrar2.click()
time.sleep(3)

pesquisa_email = driver.find_element(By.XPATH, '//*[@id="mailsearchform"]')
pesquisa_email.send_keys('Pesquisa pelo e-mail')
time.sleep(1)
pesquisa_email.send_keys(Keys.ENTER)
time.sleep(1)
pyautogui.click(517, 273, duration=1)
pyautogui.click(853, 311, duration=1)
time.sleep(8)

driver.quit()

# Mover planilha

# Lista todos os arquivos na pasta de origem
arquivos = os.listdir(pasta_origem)

# Procura pelo arquivo Excel (.xlsx) na lista de arquivos e move para a pasta de destino
for arquivo in arquivos:
    if arquivo.lower().endswith('.xlsx'):
        caminho_arquivo = os.path.join(pasta_origem, arquivo)

        # Move o arquivo para a pasta de destino
        shutil.move(caminho_arquivo, pasta_destino)

        break  # Para de procurar após encontrar o primeiro arquivo Excel

time.sleep(3)

# Lista todos os arquivos na pasta de origem
n_arquivos = os.listdir(pasta_destino)

# Contador para acompanhar o número de arquivos .xlsx encontrados
contador_xlsx = 0

# Procura pelo arquivo Excel (.xlsx) na lista de arquivos
for n_arquivo in n_arquivos:
    if n_arquivo.lower().endswith('.xlsx'):
        contador_xlsx += 1  # Incrementa o contador

#         # Verifica se é o terceiro arquivo .xlsx encontrado
        if contador_xlsx == 3:
            caminho2_arquivo = os.path.join(pasta_destino, n_arquivo)

#              # Novo nome para o arquivo Excel
            novo_nome = 'Defina o novo nome desejado'  
            novo_caminho_arquivo = os.path.join(pasta_destino, novo_nome)
            os.rename(caminho2_arquivo, novo_caminho_arquivo)

            break  # Para de procurar após encontrar o primeiro arquivo Excel


time.sleep(1)


#  Lendo a planilha cargas ontem
cargas_SR_df = pd.read_excel("cargas_sem_recebedor.xlsx")


# Multiplica os valores por 0.01 para mover duas vírgulas para a esquerda
cargas_SR_df['PESO'] = pd.to_numeric(cargas_SR_df['PESO'], errors='coerce')
cargas_SR_df['PESO'] = cargas_SR_df['PESO'] * 0.01

# Filtrar por peso acima de 1kg

colunas_selecionadas = ["AWB", "DATA COLETA", "NFS/DOC-S"]
indice_coluna_especifica = 15
coluna_especifica = cargas_SR_df.columns[indice_coluna_especifica]
colunas_selecionadas.append(coluna_especifica)
cargas_refri_df = cargas_SR_df.loc[cargas_SR_df['PESO']
                                   >= 1, colunas_selecionadas]

# Filtrando por Estado
indice_coluna_especifica = 3
caracteres_especificos = "- AP|- AM|- MS|- BA|- SE|- SC|- GO|- PA|- PI"
coluna_especifica = cargas_refri_df.columns[indice_coluna_especifica]
linhas_filtradas_df = cargas_refri_df[cargas_refri_df[coluna_especifica].str.contains(
    caracteres_especificos)]

linhas_filtradas_df['NFS/DOC-S'] = linhas_filtradas_df['NFS/DOC-S'].astype(str)

# Dividir os valores da coluna específica e replicar os dados das outras colunas
new_rows = []
for index, row in linhas_filtradas_df.iterrows():
    values = row['NFS/DOC-S'].split(',')
    for value in values:
        new_row = row.copy()
        new_row['NFS/DOC-S'] = value
        new_rows.append(new_row)

new_df = pd.DataFrame(new_rows)

new_df['NFS/DOC-S'] = new_df['NFS/DOC-S'].astype(float)

# Lendo planilha 
conservacao_df = pd.read_excel("Planilha de Conservação Luty Log.xlsx")
conservacao_df.rename(columns={'Nota Fiscal': 'NFS/DOC-S'}, inplace=True)

# Mesclar planilhas
merged_df = pd.merge(new_df, conservacao_df, on='NFS/DOC-S', how='left')
coluna_comum = 'NFS/DOC-S'
merged_df[coluna_comum] = merged_df.apply(lambda row: row[coluna_comum] if not pd.isna(
    row['NFS/DOC-S']) else row[coluna_comum], axis=1)

# Retirando os não refrigerados
coluna_verificar = 'Possui Refrigerado?'
dado_especifico = 'N'
merged_df = merged_df[merged_df[coluna_verificar] != dado_especifico]


dados_update = merged_df

# Prepara os dados

dados_update['AWB'] = dados_update['AWB'].fillna("")
dados_update['Embalagem'] = dados_update['Embalagem'].fillna("")
dados_update['INFO ADICIONAL'] = dados_update.apply(
    lambda row: '72h' if row['Embalagem'] == 'CAIXA EPS 12 LTS PD C/ GELO' else '48h', axis=1)
dados_update['PROX MANUT'] = pd.to_datetime(
    dados_update['DATA COLETA'], format='%d/%m/%Y', errors='coerce')
dados_update.loc[dados_update['INFO ADICIONAL'] ==
                 '48h', 'PROX MANUT'] += pd.Timedelta(days=2)
dados_update.loc[dados_update['INFO ADICIONAL'] ==
                 '72h', 'PROX MANUT'] += pd.Timedelta(days=3)
dados_update['PROX MANUT'] = dados_update['PROX MANUT'].fillna("").astype(str)
dados_update['PROX MANUT'] = pd.to_datetime(
    dados_update['PROX MANUT'], errors='coerce').dt.strftime('%d/%m/%Y')

dados_update['DATA COLETA'] = dados_update['DATA COLETA'].fillna(
    "").astype(str)
dados_update['DATA COLETA'] = pd.to_datetime(
    dados_update['DATA COLETA'], errors='coerce').dt.strftime('%d/%m/%Y')

dados_update['EMPRESA'] = ''
dados_update.loc[dados_update['Embalagem'].isin(
    ['CAIXA EPS 12 LTS PD C/ GELO', 'CAIXA EPS 12 LTS PS C/ GELO', 'CAIXA EPS 44 LTS PS C/ GELO', 'CAIXA EPS 44 LTS PD C/ GELO']), 'EMPRESA'] = '4BIO'

dados_update['NFS/DOC-S'] = dados_update['NFS/DOC-S'].replace([np.nan, np.inf, -np.inf], 0)

dados_update['NFS/DOC-S'] = dados_update['NFS/DOC-S'].astype(int)

print(dados_update)

# Lendo planilha cargas entregues
cargas_entregue_df = pd.read_excel("cargas_entregues.xlsx")

# Multiplica os valores por 0.01 para mover duas vírgulas para a esquerda
cargas_entregue_df['PESO'] = pd.to_numeric(
    cargas_entregue_df['PESO'], errors='coerce')
cargas_entregue_df['PESO'] = cargas_entregue_df['PESO'] * 0.01

# Filtrar por peso acima de 1kg
col_selecionadas2 = ["NFS/DOC-S"]
ind_coluna_especifica2 = 15
col_selecionadas2.append(cargas_entregue_df.columns[ind_coluna_especifica2])
cargas_refri_df2 = cargas_entregue_df.loc[cargas_entregue_df['PESO']
                                          >= 1, col_selecionadas2]

# Filtrando por Estado
indice_coluna_especifica3 = 1
caracteres_especificos2 = "- AP|- AM|- MS|- BA|- SE|- SC|- GO|- PA|- PI"
coluna_especifica2 = cargas_refri_df2.columns[indice_coluna_especifica3]
linhas_filtradas_df2 = cargas_refri_df2[cargas_refri_df2[coluna_especifica2].str.contains(
    caracteres_especificos2)]


nfs_doc_s_df_unique = linhas_filtradas_df2.iloc[:, 0].unique()
nfs_doc_s_df_unique = [str(item).strip() for item in nfs_doc_s_df_unique]


# Adiciona todos os escopos ao acesso à planilha
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']


# O ID e o intervalo de uma planilha de amostra.
SAMPLE_SPREADSHEET_ID = 'ID da planilha google sheets'
SHEET_NAME = 'CONTROLE'


# Fazer login no google sheets
def main():
    creds = None

    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    # Alterações na planilha
    try:
        service = build('sheets', 'v4', credentials=creds)

        sheet = service.spreadsheets()

        # Busca qual a próxima linha vazia para preenchimento
        coluna = "C"

        resultado = service.spreadsheets().values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                                        range=f'CONTROLE!{coluna}:{coluna}').execute()
        val = resultado.get('values', [])
        proxima = len(val) + 1
        proxima_linha = f"B{proxima}"
        print(proxima_linha)

        # Alimenta os dados
        colunas_escolhidas = ['AWB', 'DATA COLETA', 'NFS/DOC-S',
                              'CIDADE.1', 'INFO ADICIONAL', 'PROX MANUT', 'EMPRESA']

        valores_adicionar = dados_update[colunas_escolhidas].values.tolist()
        print(valores_adicionar)

        result = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                       range=proxima_linha, valueInputOption="USER_ENTERED", body={'values': valores_adicionar}).execute()

        # PESQUISAR NA PLANILHA AS LINHAS QUE CONTEM AS NFS JÁ ENTREGUES E SALVAR A INFORMAÇÃO DA LINHA
        spreadsheet_id = SAMPLE_SPREADSHEET_ID
        range_name = f"{SHEET_NAME}!D2:D"

        resultad = service.spreadsheets().values().get(
            spreadsheetId=spreadsheet_id, range=range_name).execute()
        nfs_doc_s_planilha = resultad.get('values', [])
        nfs_doc_s_planilha = [
            item for sublist in nfs_doc_s_planilha for item in sublist]

        valores_comuns = list(set(nfs_doc_s_planilha) &
                              set(nfs_doc_s_df_unique))

        # Selecionar os primeiros 60 valores
        primeiros_60 = valores_comuns[:60]

        # Selecionar os segundos 60 valores
        segundos_60 = valores_comuns[60:120]

        # Selecionar o restante dos valores
        restante = valores_comuns[120:]

        print(f"primeiros 60: ", primeiros_60)
        print(f"segundos 60: ", segundos_60)
        print(f"restante: ", restante)

        sheet_id = 0

        for valores in [primeiros_60, segundos_60, restante]:
            for valor_comum in valores:
                # Encontre o índice da linha que contém o valor_comum na coluna NFS/DOC-S
                linha_encontrada = None
                for index, valor_planilha in enumerate(nfs_doc_s_planilha, start=2):
                    if valor_planilha == valor_comum:
                        linha_encontrada = index
                        break

                if linha_encontrada:
                    # Defina a faixa de colunas a ser formatada
                    range_notation = f'A{linha_encontrada}:Z{linha_encontrada}'
                    color = {"red": 0.5, "green": 0.5,
                             "blue": 0.5}  # Cor cinza
                    body = {
                        "requests": [
                            {
                                "repeatCell": {
                                    "range": {"sheetId": sheet_id, "startRowIndex": linha_encontrada - 1, "endRowIndex": linha_encontrada},
                                    "cell": {"userEnteredFormat": {"backgroundColor": color}},
                                    "fields": "userEnteredFormat.backgroundColor",
                                }
                            }
                        ]
                    }

                    response = service.spreadsheets().batchUpdate(
                        spreadsheetId=SAMPLE_SPREADSHEET_ID, body=body).execute()
            time.sleep(60)

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()

time.sleep(0.5)
# Excluir arquivos Excel na pasta
current_directory = os.getcwd()
excel_files = [f for f in os.listdir(
    current_directory) if f.endswith('.xlsx') or f.endswith('.xls')]

for excel_file in excel_files:
    try:
        os.remove(excel_file)
        print(f"Arquivo {excel_file} excluído com sucesso.")
    except Exception as e:
        print(f"Erro ao excluir {excel_file}: {e}")