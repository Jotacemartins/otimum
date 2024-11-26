from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
from datetime import datetime
import time

# Caminho para o driver do navegador
DRIVER_PATH = 'caminho/do/chromedriver'

# Configurar Selenium
options = webdriver.ChromeOptions()
options.add_argument('--start-maximized')
driver = webdriver.Chrome(executable_path=DRIVER_PATH, options=options)

# Abrir o site
driver.get('URL_DO_SITE_DE_CONSULTA')

# Planilha de dados
planilha = 'dados_cartoes.xlsx'
try:
    df = pd.read_excel(planilha)
except FileNotFoundError:
    # Criar planilha se não existir
    df = pd.DataFrame(columns=['Nome', 'Cartão', 'Saldo', 'Data'])

# Lista de dados (nome e número do cartão)
cartoes = [
    {'nome': 'Fulano', 'cartao': '12345678901234'},
    {'nome': 'Ciclano', 'cartao': '98765432109876'}
]

for pessoa in cartoes:
    nome = pessoa['nome']
    numero_cartao = pessoa['cartao']

    # Interagir com os elementos
    campo_cartao = driver.find_element(By.ID, 'ID_DO_CAMPO_DO_CARTAO')
    campo_cartao.clear()
    campo_cartao.send_keys(numero_cartao)
    
    botao_consultar = driver.find_element(By.ID, 'ID_DO_BOTAO_CONSULTAR')
    botao_consultar.click()
    
    time.sleep(3)  # Esperar a resposta

    # Capturar informações
    saldo = driver.find_element(By.ID, 'ID_DO_CAMPO_SALDO').text
    data_consulta = datetime.now().strftime('%d/%m/%Y %H:%M')

    # Tirar print
    driver.save_screenshot(f'print_{numero_cartao}.png')

    # Atualizar planilha
    df = df.append({'Nome': nome, 'Cartão': numero_cartao, 'Saldo': saldo, 'Data': data_consulta}, ignore_index=True)

# Salvar planilha
df.to_excel(planilha, index=False)

# Fechar o navegador
driver.quit()
