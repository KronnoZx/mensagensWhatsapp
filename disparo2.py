import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time
import pyautogui as pya
import urllib
import os

navegador = webdriver.Chrome()
navegador.get("https://web.whatsapp.com")

# zap carregador
tabela = pd.read_excel("base.xlsx")

# para cada linha da planilha
for linha in tabela.index:
    # localiza informações na planilha
    nome = tabela.loc[linha, "nome"]
    mensagem = tabela.loc[linha, "mensagem"]
    arquivo = tabela.loc[linha, "arquivo"]
    telefone = tabela.loc[linha, "telefone"]
   # enviarArquivo = 

    # troca nome pela coluna nome e converte para linguagem de url sei la
    texto = mensagem.replace("$nome", nome)
    texto = urllib.parse.quote(texto)

    # cria o link com o telefone e mensagem
    link = f"https://web.whatsapp.com/send?phone={telefone}&text={texto}"

    navegador.get(link)
    # espera o zap carregar
    while len(navegador.find_elements('id', 'side')) < 1:
        time.sleep(1)
    time.sleep(2)  # só pra garantir que o zap vai carregar

    # verifica se o número é válido
    if len(navegador.find_elements('xpath', '//*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) < 1:
        # envia a mensagem
        navegador.find_element(
            'xpath', '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
        if arquivo != "N":
            caminho_completo = os.path.abspath(f"arquivos/{arquivo}")
            print(caminho_completo)
            navegador.find_element(
                'xpath', '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/span').click()  # clica no clips
            time.sleep(2)
            navegador.find_element(
                'xpath', '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[1]/button/input').send_keys(caminho_completo)
            time.sleep(3)
            navegador.find_element(
                'xpath', '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div').click()

            time.sleep(3)