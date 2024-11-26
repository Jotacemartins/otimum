import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import pandas as pd
from datetime import datetime
import time

# Caminho para o ChromeDriver
DRIVER_PATH = './chromedriver.exe'

# Configurações do navegador
options = webdriver.ChromeOptions()
options.add_argument('--start-maximized')  # Iniciar o navegador maximizado
options.add_argument('--disable-gpu')      # Melhorar compatibilidade


def processar_arquivo(file_path):
    try:
        service = Service(DRIVER_PATH)
        driver = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao iniciar o ChromeDriver: {e}")
        return

    SITE_URL = 'https://example.com/consulta-saldo'  # Atualize com a URL real
    planilha = 'dados_cartoes.xlsx'

    try:
        df = pd.read_excel(planilha)
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Nome', 'Cartão', 'Saldo', 'Data'])

    try:
        with open(file_path, 'r') as file:
            linhas = file.readlines()
            cartoes = [{'nome': linha.split(',')[0].strip(), 'cartao': linha.split(',')[1].strip()} for linha in linhas]
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo: {e}")
        driver.quit()
        return

    for pessoa in cartoes:
        nome = pessoa['nome']
        numero_cartao = pessoa['cartao']
        print(f"Iniciando consulta para {numero_cartao}...")

        try:
            # Acessar o site
            driver.get(SITE_URL)

            # Garantir que a página carregue completamente
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'page_input-field__61j_k'))
            )

            # Localizar e preencher o campo do cartão
            campo_cartao = driver.find_element(By.CLASS_NAME, 'page_input-field__61j_k')
            campo_cartao.clear()
            campo_cartao.send_keys(numero_cartao)
            print(f"Número do cartão inserido: {numero_cartao}")

            # Localizar e clicar no botão
            botao_consultar = driver.find_element(By.CLASS_NAME, 'page_submit-button__r0H3S')
            botao_consultar.click()
            print("Botão 'Consultar saldo' clicado.")

            # Aguardar a resposta
            container_info = WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'page_card-info__gM7vJ'))
            )
            print("Informações do cartão carregadas.")

            # Capturar saldo e data
            saldo = container_info.find_element(By.CLASS_NAME, 'page_saldo__5B0iu').text
            saldo = saldo.replace('Saldo: R$ ', '').strip()

            data_consulta = datetime.now().strftime('%d/%m/%Y %H:%M')

            # Salvar na planilha
            df = df.append({'Nome': nome, 'Cartão': numero_cartao, 'Saldo': saldo, 'Data': data_consulta}, ignore_index=True)

            # Salvar print da página
            nome_arquivo_print = f'print_{numero_cartao}.png'
            driver.save_screenshot(nome_arquivo_print)
            print(f"Consulta concluída para {numero_cartao}. Saldo: {saldo}. Print salvo: {nome_arquivo_print}")

        except TimeoutException as e:
            print(f"Erro de tempo ao consultar o cartão {numero_cartao}: {e}")
        except NoSuchElementException as e:
            print(f"Elemento necessário não encontrado para o cartão {numero_cartao}: {e}")
        except Exception as e:
            print(f"Erro inesperado ao processar {numero_cartao}: {e}")

        # Delay entre consultas
        time.sleep(10)

    # Salvar a planilha
    df.to_excel(planilha, index=False)
    print(f"Dados salvos na planilha: {planilha}")

    driver.quit()
    messagebox.showinfo("Sucesso", f"Consultas concluídas e dados salvos na planilha: {planilha}")


def selecionar_arquivo():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo TXT",
        filetypes=(("Arquivos de Texto", "*.txt"), ("Todos os Arquivos", "*.*"))
    )
    if file_path:
        processar_arquivo(file_path)


root = tk.Tk()
root.title("Consulta de Saldo de Cartões")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

btn_selecionar = tk.Button(frame, text="Selecionar Arquivo TXT", command=selecionar_arquivo, padx=10, pady=5)
btn_selecionar.pack()

root.mainloop()
