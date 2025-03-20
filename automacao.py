import string
import traceback

import pyautogui
from selenium.webdriver.common.by import By
import undetected_chromedriver as uc
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import datetime
import pandas as pd

import pygetwindow as gw
import time
import os


def posicionar_prompt_direita():
    # Define o título do CMD para facilitar a busca
    os.system("title Automacao")
    time.sleep(1)  # Aguarda um tempo para o CMD processar a alteração

    try:
        # Obtém a janela do prompt de comando
        janela = gw.getWindowsWithTitle("Automacao")[0]

        # Obtém o tamanho da tela
        screen_width, screen_height = pyautogui.size()

        # Define nova posição (lado direito da tela)
        largura_janela = 600  # Ajuste conforme necessário
        altura_janela = 400  # Ajuste conforme necessário
        nova_x = screen_width - largura_janela
        nova_y = 50  # Mantém um pequeno espaço do topo

        # Define tamanho e posição da janela
        janela.moveTo(nova_x, nova_y)
        janela.resizeTo(largura_janela, altura_janela)

        # Mantém sempre no topo
        janela.alwaysOnTop = True
    except IndexError:
        print("Não foi possível encontrar a janela do Prompt de Comando.")

posicionar_prompt_direita()


# Arquivo com lista de cnpj
arquivo = open("lista.txt", "r")
data = arquivo.read()
cnpj_list = data.split("\n")
arquivo.close()


def iniciar_navegador():
    options = uc.ChromeOptions()
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--start-maximized")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-popup-blocking")

    # Inicializa o Chrome usando undetected-chromedriver
    driver = uc.Chrome(options=options, headless=False)
    driver.set_window_size(800, 600)
    return driver


def consultar_cnpj(driver, cnpj):

    # Acessa o site e espera até carregar
    driver.get("https://www8.receita.fazenda.gov.br/simplesnacional/aplicacoes.aspx?id=21")
    WebDriverWait(driver, 999).until(EC.presence_of_element_located((By.ID, "frame")))
    iframe_element = driver.find_element(By.ID, "frame")
    iframe_posicao = iframe_element.location
    driver.switch_to.frame(driver.find_element(By.ID, "frame"))
    try:
        # # Aguarda o campo de CNPJ estar presente
        WebDriverWait(driver, 999).until(EC.presence_of_element_located((By.ID, "Cnpj")))



        # # Insere o CNPJ]
        campo_cnpj = driver.find_element(By.ID, "Cnpj")
        campo_cnpj.clear()
        campo_cnpj.send_keys(cnpj)

        btn_mais_info = None
        contador = 0
        while not btn_mais_info:
            contador += 1
            if contador == 5:
                contador = 0

            # pega a posição relativa do botão Consultar
            consultar_button = driver.find_element(By.CLASS_NAME, "btn")
            x = consultar_button.location['x'] + consultar_button.size['width'] / 2
            y = consultar_button.location['y'] + consultar_button.size['height'] / 2 + 100

            # Converter as coordenadas para o contexto da página principal
            centro_x_pagina_principal = x + iframe_posicao['x']
            centro_y_pagina_principal = y + iframe_posicao['y']

            # Move para o botão de consultar
            pyautogui.moveTo(centro_x_pagina_principal + 30, centro_y_pagina_principal, duration=0.2)

            # Clica no botão consultar
            print("Provocando erro 'Comportamento de robô'")
            for i in range(3):
                pyautogui.click()
            time.sleep(0.5)

            text_error = driver.find_elements(By.CLASS_NAME, "text-danger")
            print("Testando validade do CNPJ")
            if text_error:
                if text_error[0].size['width'] > 0:
                    raise Exception("Cnpj invalido")

            WebDriverWait(driver, timeout=5).until(EC.presence_of_element_located((By.CLASS_NAME, "alert")))
            consultar_button = driver.find_element(By.CLASS_NAME, "btn")
            x = consultar_button.location['x'] + consultar_button.size['width'] / 2 + (contador * 8)
            y = consultar_button.location['y'] + consultar_button.size['height'] / 2 + 100  # ajuste para barra do navegador
            centro_x_pagina_principal = x + iframe_posicao['x']
            centro_y_pagina_principal = y + iframe_posicao['y']
            driver.switch_to.default_content()

            # Move para a nova posição do botão após aparecer a msg de comportamento de robô
            pyautogui.moveTo(centro_x_pagina_principal -10, centro_y_pagina_principal, duration=0.2)

            # Clica no botão consultar na nova posição
            print("Clicando no botão na nova posição para e esperando a próxima aba")
            for i in range(3):
                pyautogui.click()
                time.sleep(0.5)

            driver.switch_to.frame(driver.find_element(By.ID, "frame"))
            try:
                WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "btnMaisInfo")))
            except Exception as e:
                print(f"Erro {e} {traceback.format_exc()}")
            finally:
                btn_mais_info = driver.find_elements(By.ID, "btnMaisInfo")

        # Clica em "Mais informações"
        print("Clicando em 'Mais Informações'")
        driver.find_element(By.ID, "btnMaisInfo").click()
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, "in")))

        # Encontrar os elementos necessários para a tabela
        print("Pegando dados necessários")
        sit_atual = driver.find_elements(By.CLASS_NAME, "panel-body")[1]
        sit_atual_simples = sit_atual.find_elements(By.TAG_NAME, "span")[0].text
        sit_atual_simei = sit_atual.find_elements(By.TAG_NAME, "span")[1].text
        mais_info = driver.find_element(By.ID, "maisInfo")
        periodos_anteriores_simples = None
        periodos_anteriores_simei = None
        # Pega a primeira linha das tabelas de periodos anteriores
        tabela = mais_info.find_elements(By.CLASS_NAME, "table")
        if len(tabela) != 0:
            linhas_simples = tabela[0].find_elements(By.TAG_NAME, "tr")
            colunas = linhas_simples[1].find_elements(By.TAG_NAME, "td")
            for coluna in colunas:
                if periodos_anteriores_simples is not None:
                    periodos_anteriores_simples += "-" + coluna.text
                else:
                    periodos_anteriores_simples = coluna.text
        else:
            periodos_anteriores_simples = "Sem Dados"
        if len(tabela) == 2:
            linhas_simei = tabela[1].find_elements(By.TAG_NAME, "tr")
            colunas = linhas_simei[1].find_elements(By.TAG_NAME, "td")
            for coluna in colunas:
                if periodos_anteriores_simei is not None:
                    periodos_anteriores_simei += "-" + coluna.text
                else:
                    periodos_anteriores_simei = coluna.text
        else:
            periodos_anteriores_simei = "Sem Dados"
        # Pega a data atual
        data_da_consulta = datetime.datetime.today().strftime('%d-%m-%Y')
        # Salva tudo e retorna pronto para salvar em planilha
        dado = [data_da_consulta, cnpj , sit_atual_simples, sit_atual_simei, periodos_anteriores_simples, periodos_anteriores_simei]
        print(dado)
        return dado
    except Exception as e:
        print(f"Erro ao consultar CNPJ {cnpj}: {e} {traceback.format_exc()}")
        data_da_consulta = datetime.datetime.today().strftime('%d-%m-%Y')
        dado = [data_da_consulta, cnpj, "ERRO", "ERRO", "ERRO", "ERRO"]
        return dado


def main():
    data_inicio = datetime.datetime.now()
    contador = 0
    driver = iniciar_navegador()
    col = ["Data da Consulta", "CNPJ", "Situacao atual simples", "Situacao atual Simei", "Periodos Anteriores Simples", "Periodos Anteriores Simei"]
    dados = pd.DataFrame(columns = col)
    cnpj_atual = None
    try:
        for cnpj in cnpj_list:
            contador += 1
            cnpj_atual = cnpj
            cnpj = cnpj.translate(str.maketrans("", "", string.punctuation))
            print(f"{contador} - Consultando CNPJ: {cnpj}")
            print(f"CNPJ {contador} de {len(cnpj_list)}")
            dado = consultar_cnpj(driver,cnpj)
            dados.loc[cnpj] = dado

    except Exception as e:
        print("Erro no cnpj: " + cnpj_atual)
        print(f"Erro {e} {traceback.format_exc()}")
    finally:
        print(dados)
        dados.to_excel(f"dados_simples_simei_{datetime.datetime.today().strftime('%Y_%m_%d--%H_%M_%S')}.xlsx", index = False)
        print(f"O inicio foi: {data_inicio}")
        print(f"O final foi: {datetime.datetime.now()}")
        print(f"Foram consultados {contador} cnpj's")
        input("Pressione ENTER para continuar")
        quit()

if __name__ == "__main__":
    main()
