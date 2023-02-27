import selenium
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support import expected_conditions
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import pyautogui as pya
import urllib
import os

navegador = webdriver.Chrome()
navegador.get("https://web.whatsapp.com")

tabela = pd.read_excel("baseVers1.1.xlsx")
print(tabela)

for linha in tabela.index:

    mensagem = tabela.loc[linha, 'mensagem']
    arquivo = tabela.loc[linha, 'arquivo']
    telefone = tabela.loc[linha, 'telefone']

    texto = urllib.parse.quote(mensagem)

    if mensagem != "N":
        link = f"https://web.whatsapp.com/send?phone={telefone}&text={texto}"
    else:
        link = f"https://web.whatsapp.com/send?phone={telefone}"

    navegador.get(link)

    while len(navegador.find_elements('id', 'side')) < 1:
        time.sleep(1)
    time.sleep(2)

    if len(navegador.find_elements('xpath', '/*[@id="app"]/div/span[2]/div/span/div/div/div/div/div/div[1]')) < 1:

        navegador.find_element(
            'xpath', '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button/span').click()
        if arquivo != "N":
            caminho_completo = os.path.abspath(f'arquivos/{arquivo}')
            if mensagem != "N":
                navegador.find_element(
                    'xpath', '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/span').click()
            else:
                print("enviado")
            time.sleep(3)
            navegador.find_element(
                'xpath', '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/span/div/div/ul/li[1]/button/input').send_keys(caminho_completo)
            time.sleep(3)
            navegador.find_element(
                'xpath', '//*[@id="app"]/div/div/div[2]/div[2]/span/div/span/div/div/div[2]/div/div[2]/div[2]/div/div').click()
            time.sleep(3)
        else:
            navegador.find_element(
                'xpath', '//*[@id="main"]/footer/div[1]/div/span[2]/div/div[1]/div[2]/div/div/span').click()
            time.sleep(3)
