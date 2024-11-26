import tkinter as tk
from tkinter import filedialog, messagebox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service
import pandas as pd
from datetime import datetime
import time

# Caminho para o ChromeDriver (na mesma pasta do código)
DRIVER_PATH = './chromedriver.exe'

# Configurações do navegador
options = webdriver.ChromeOptions()
options.add_argument('--start-maximized')  # Iniciar o navegador maximizado
options.add_argument('--disable-gpu')      # Melhorar compatibilidade
# options.add_argument('--headless')        # Opcional: executa sem interface gráfica


def processar_arquivo(file_path):
    # Inicializar o driver
    try:
        service = Service(DRIVER_PATH)
        driver = webdriver.Chrome(service=service, options=options)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao iniciar o ChromeDriver: {e}")
        return

    # URL do site de consulta (substitua pela URL correta)
    SITE_URL = 'https://example.com/consulta-saldo'

    # Planilha para salvar os dados
    planilha = 'dados_cartoes.xlsx'

    # Abrir a planilha existente ou criar uma nova
    try:
        df = pd.read_excel(planilha)
    except FileNotFoundError:
        df = pd.DataFrame(columns=['Nome', 'Cartão', 'Saldo', 'Data'])

    # Ler os dados do arquivo TXT
    try:
        with open(file_path, 'r') as file:
            linhas = file.readlines()
            cartoes = [{'nome': linha.split(',')[0].strip(), 'cartao': linha.split(',')[1].strip()} for linha in linhas]
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao ler o arquivo: {e}")
        driver.quit()
        return

    # Abrir o site
    try:
        driver.get(SITE_URL)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao acessar o site: {e}")
        driver.quit()
        return

    # Loop para cada cartão
    for pessoa in cartoes:
        nome = pessoa['nome']
        numero_cartao = pessoa['cartao']
        print(f"Consultando saldo para o cartão {numero_cartao}...")

        try:
            # Localizar o campo do cartão usando a classe (campo_input)
            campo_cartao = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.CLASS_NAME, 'page_input-field__61j_k'))  # Usando a classe
            )
            campo_cartao.clear()
            campo_cartao.send_keys(numero_cartao)

            # Aguardar o botão de consulta e clicar nele
            botao_consultar = WebDriverWait(driver, 15).until(
                EC.element_to_be_clickable((By.CLASS_NAME, 'page_submit-button__r0H3S'))  # Usando a classe
            )
            botao_consultar.click()

            # Aguardar o saldo ser exibido
            saldo = WebDriverWait(driver, 15).until(
                EC.presence_of_element_located((By.ID, 'saldo_id'))  # Substitua 'saldo_id' com o ID real do saldo
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

        # Delay de 10 segundos entre as consultas
        time.sleep(10)

    # Salvar a planilha atualizada
    df.to_excel(planilha, index=False)
    messagebox.showinfo("Sucesso", f"Dados salvos na planilha: {planilha}")

    # Fechar o navegador
    driver.quit()


def selecionar_arquivo():
    # Abrir janela para selecionar arquivo
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo TXT",
        filetypes=(("Arquivos de Texto", "*.txt"), ("Todos os Arquivos", "*.*"))
    )
    if file_path:
        processar_arquivo(file_path)


# Interface gráfica
root = tk.Tk()
root.title("Consulta de Saldo de Cartões")

# Configurar a janela
frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

# Botão para selecionar o arquivo
btn_selecionar = tk.Button(frame, text="Selecionar Arquivo TXT", command=selecionar_arquivo, padx=10, pady=5)
btn_selecionar.pack()

# Rodar a interface
root.mainloop()
