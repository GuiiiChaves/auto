import os
import time
import pandas as pd
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager 
from selenium.webdriver.chrome.service import Service

# Configurando o serviço do driver do Chrome
servico = Service(ChromeDriverManager().install())

# Inicializando o navegador
navegador = webdriver.Chrome(service=servico)

# Abrindo a página
navegador.get('https://www.rpachallenge.com/')

# Esperando um pouco para a página carregar completamente
time.sleep(2)

# Clicando no link para baixar o arquivo
navegador.find_element('xpath', '/html/body/app-root/div[2]/app-rpa1/div/div[1]/div[6]/a').click()

# Esperando o download ser concluído (você pode ajustar o tempo dependendo do tamanho do arquivo)
time.sleep(5)

# Fechando o navegador, pois não precisaremos mais dele
navegador.quit()

# Abrindo o diretório de downloads
diretorio_downloads = r'C:\Users\guilh\Downloads'  # Substitua pelo caminho real do diretório de downloads

# Obtendo a lista de arquivos no diretório de downloads
arquivos = os.listdir(diretorio_downloads)

# Filtrando apenas arquivos com extensão .xlsx
arquivos_xlsx = [arquivo for arquivo in arquivos if arquivo.endswith('.xlsx')]

# Verificando se existem arquivos .xlsx baixados
if arquivos_xlsx:
    # Abrindo o arquivo Excel mais recente
    caminho_arquivo = os.path.join(diretorio_downloads, arquivos_xlsx[0])
    df = pd.read_excel(caminho_arquivo)
    
    # Convertendo os dados do DataFrame para uma lista de dicionários
    dados = df.to_dict('records')
    
    # Exibindo os dados
    print(dados)
else:
    print("Nenhum arquivo .xlsx foi encontrado no diretório de downloads.")
