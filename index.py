from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
from datetime import datetime

# Caminho para o driver na mesma pasta do código
DRIVER_PATH = './chromedriver.exe'

# Configurações do navegador
options = webdriver.ChromeOptions()
options.add_argument('--start-maximized')  # Iniciar o navegador maximizado
options.add_argument('--disable-gpu')      # Melhorar compatibilidade
options.add_argument('--disable-extensions')  # Desativar extensões
options.add_argument('--headless')        # Opcional: executa sem interface gráfica

# Inicializar o driver
try:
    driver = webdriver.Chrome(executable_path=DRIVER_PATH, options=options)
except Exception as e:
    print("Erro ao iniciar o ChromeDriver:", e)
    exit()

# URL do site de consulta (substitua pela URL correta)
SITE_URL = 'https://example.com/consulta-saldo'

# Planilha para salvar os dados
planilha = 'dados_cartoes.xlsx'

# Abrir a planilha existente ou criar uma nova
try:
    df = pd.read_excel(planilha)
except FileNotFoundError:
    df = pd.DataFrame(columns=['Nome', 'Cartão', 'Saldo', 'Data'])

# Lista de cartões para consulta
cartoes = [
    {'nome': 'Fulano', 'cartao': '12345678901234'},
    {'nome': 'Ciclano', 'cartao': '98765432109876'}
]

# Abrir o site
try:
    driver.get(SITE_URL)
    print("Site acessado com sucesso.")
except Exception as e:
    print("Erro ao acessar o site:", e)
    driver.quit()
    exit()

# Loop para cada cartão
for pessoa in cartoes:
    nome = pessoa['nome']
    numero_cartao = pessoa['cartao']
    print(f"Consultando saldo para o cartão {numero_cartao}...")

    try:
        # Localizar o campo do cartão (substitua pelo ID correto do campo no site)
        campo_cartao = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'campo_cartao_id'))  # Substitua 'campo_cartao_id'
        )
        campo_cartao.clear()
        campo_cartao.send_keys(numero_cartao)
        
        # Clicar no botão de consulta (substitua pelo ID correto do botão no site)
        botao_consultar = driver.find_element(By.ID, 'botao_consultar_id')  # Substitua 'botao_consultar_id'
        botao_consultar.click()
        
        # Aguardar o saldo ser exibido
        saldo = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'saldo_id'))  # Substitua 'saldo_id'
        ).text

        # Capturar data da consulta
        data_consulta = datetime.now().strftime('%d/%m/%Y %H:%M')

        # Salvar dados na planilha
        df = df.append({'Nome': nome, 'Cartão': numero_cartao, 'Saldo': saldo, 'Data': data_consulta}, ignore_index=True)

        # Capturar um print da tela
        driver.save_screenshot(f'print_{numero_cartao}.png')

        print(f"Consulta realizada com sucesso para o cartão {numero_cartao}. Saldo: {saldo}")

    except TimeoutException:
        print(f"Erro: Tempo excedido para consultar o cartão {numero_cartao}.")
    except NoSuchElementException:
        print(f"Erro: Elementos necessários não encontrados para o cartão {numero_cartao}.")
    except Exception as e:
        print(f"Erro inesperado para o cartão {numero_cartao}: {e}")

# Salvar a planilha atualizada
df.to_excel(planilha, index=False)
print(f"Dados salvos na planilha: {planilha}")

# Fechar o navegador
driver.quit()
