from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
from datetime import datetime
import time

# Caminho para o driver do navegador (atualize para o local correto)
DRIVER_PATH = 'C:/Users/SeuUsuario/Downloads/chromedriver.exe'

# Configurações do Selenium
options = webdriver.ChromeOptions()
options.add_argument('--start-maximized')
options.add_argument('--disable-gpu')  # Para melhorar compatibilidade
options.add_argument('--headless')    # Executa sem abrir o navegador (opcional)

# Iniciar o driver
driver = webdriver.Chrome(executable_path=DRIVER_PATH, options=options)

# URL do site de consulta (substitua pela URL correta)
SITE_URL = 'https://example.com/consulta-saldo'

# Planilha de dados
planilha = 'dados_cartoes.xlsx'

# Tentar abrir a planilha existente ou criar uma nova
try:
    df = pd.read_excel(planilha)
except FileNotFoundError:
    df = pd.DataFrame(columns=['Nome', 'Cartão', 'Saldo', 'Data'])

# Lista de dados para consultar (nome e número do cartão)
cartoes = [
    {'nome': 'Fulano', 'cartao': '12345678901234'},
    {'nome': 'Ciclano', 'cartao': '98765432109876'}
]

# Abrir o site
try:
    driver.get(SITE_URL)
except Exception as e:
    print("Erro ao abrir o site:", e)
    driver.quit()
    exit()

# Loop para cada cartão
for pessoa in cartoes:
    nome = pessoa['nome']
    numero_cartao = pessoa['cartao']
    
    try:
        # Localizar o campo do cartão (substitua pelo ID ou seletor correto)
        campo_cartao = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'campo_cartao_id'))  # Substitua 'campo_cartao_id'
        )
        campo_cartao.clear()
        campo_cartao.send_keys(numero_cartao)
        
        # Clicar no botão de consultar saldo
        botao_consultar = driver.find_element(By.ID, 'botao_consultar_id')  # Substitua 'botao_consultar_id'
        botao_consultar.click()
        
        # Esperar pela exibição do saldo
        saldo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'saldo_id'))  # Substitua 'saldo_id'
        ).text

        # Registrar a data da consulta
        data_consulta = datetime.now().strftime('%d/%m/%Y %H:%M')

        # Capturar um print da tela
        driver.save_screenshot(f'print_{numero_cartao}.png')

        # Atualizar a planilha
        df = df.append({'Nome': nome, 'Cartão': numero_cartao, 'Saldo': saldo, 'Data': data_consulta}, ignore_index=True)

        print(f"Consulta realizada com sucesso para o cartão {numero_cartao}.")

    except TimeoutException:
        print(f"Tempo excedido para consultar o cartão {numero_cartao}. Verifique o site.")
    except NoSuchElementException:
        print(f"Erro ao localizar elementos para o cartão {numero_cartao}.")
    except Exception as e:
        print(f"Erro inesperado para o cartão {numero_cartao}: {e}")

# Salvar a planilha
df.to_excel(planilha, index=False)
print(f"Dados salvos na planilha {planilha}.")

# Fechar o navegador
driver.quit()
