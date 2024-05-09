import time
import os
import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager 
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By

# Configurando o serviço do driver do Chrome
servico = Service(ChromeDriverManager().install())

# Inicializando o navegador
navegador = webdriver.Chrome(service=servico)

# Abrindo a página
navegador.get('https://www.rpachallenge.com/')

# Esperando um pouco para a página carregar completamente
time.sleep(2)

# Clicando no link para baixar o arquivo
navegador.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/a').click()

# Esperando o download ser concluído (você pode ajustar o tempo dependendo do tamanho do arquivo)
time.sleep(5)

# Abrindo o diretório de downloads
diretorio_downloads = r'C:\Users\guilh\Downloads'  # Substitua pelo caminho real do diretório de downloads

# Obtendo a lista de arquivos no diretório de downloads
arquivos = os.listdir(diretorio_downloads)

# Filtrando apenas arquivos com extensão .xlsx
arquivos_xlsx = [arquivo for arquivo in arquivos if arquivo.endswith('.xlsx')]

# Verificando se existem arquivos .xlsx baixados
if arquivos_xlsx:
    # Encontrando o arquivo Excel mais recente
    arquivo_mais_recente = max(arquivos_xlsx, key=lambda x: os.path.getctime(os.path.join(diretorio_downloads, x)))
    caminho_arquivo = os.path.join(diretorio_downloads, arquivo_mais_recente)
    
    # Carregando o arquivo Excel com pandas
    dados_excel = pd.read_excel(caminho_arquivo)
    dados_usuarios = dados_excel.to_dict('records')
else:
    print("Nenhum arquivo .xlsx foi encontrado no diretório de downloads.")

# Esperando um pouco para a página carregar completamente
time.sleep(2)

# Preenchendo o formulário para cada usuário
for usuario in dados_usuarios:
    first_name = usuario.get("First Name", "")
    last_name = usuario.get("Last Name", "") 
    address = usuario.get("Address", "")
    phone_number = usuario.get("Phone Number", "")
    email = usuario.get("Email", "")
    company_name = usuario.get("Company Name", "")
    role_in_company = usuario.get("Company Role", "")
    
    # Preenchendo o formulário com os dados do usuário atual
    navegador.find_element(By.XPATH, '//input[@ng-reflect-name="labelFirstName"]').send_keys(first_name)
    navegador.find_element(By.XPATH, '//input[@ng-reflect-name="labelLastName"]').send_keys(last_name)
    navegador.find_element(By.XPATH, '//input[@ng-reflect-name="labelCompanyName"]').send_keys(company_name)
    
    # Pequeno atraso antes de preencher o campo Company Role
    time.sleep(1)
    
    navegador.find_element(By.XPATH, '//input[@ng-reflect-name="labelCompanyRole"]').send_keys(role_in_company)
    navegador.find_element(By.XPATH, '//input[@ng-reflect-name="labelAddress"]').send_keys(address)
    navegador.find_element(By.XPATH, '//input[@ng-reflect-name="labelEmail"]').send_keys(email)
    navegador.find_element(By.XPATH, '//input[@ng-reflect-name="labelPhone"]').send_keys(phone_number)
    
    # Submeter o formulário
    navegador.find_element(By.XPATH, '/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input').click()


    


