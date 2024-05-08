import time
import os
import openpyxl
import subprocess
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
    
    # Abrindo o arquivo Excel
    subprocess.Popen(['start', 'excel.exe', caminho_arquivo], shell=True)
else:
    print("Nenhum arquivo .xlsx foi encontrado no diretório de downloads.")


# Abrindo o arquivo Excel
workbook = openpyxl.load_workbook(caminho_arquivo)
sheet = workbook.active

# Lista para armazenar os dados dos usuários
dados_usuarios = []

# Lendo os dados da planilha linha por linha (ignorando o cabeçalho)
for row in sheet.iter_rows(min_row=2, values_only=True):
    usuario = {
        "First Name": row[0],
        "Last Name": row[1],
        "Address": row[2],
        "City": row[3],
        "State": row[4],
        "Zip": row[5],
        "Phone Number": row[6],
        "Email": row[7]
    }
    dados_usuarios.append(usuario)

# Fechando o navegador
navegador.quit()

# Abrindo o site
navegador.get("https://www.rpachallenge.com/")

# Esperando um pouco para a página carregar completamente
time.sleep(2)

# Preenchendo o formulário para cada usuário
for usuario in dados_usuarios:
    first_name = usuario["First Name"]
    last_name = usuario["Last Name"]
    address = usuario["Address"]
    city = usuario["City"]
    state = usuario["State"]
    zip_code = usuario["Zip"]
    phone_number = usuario["Phone Number"]
    email = usuario["Email"]
    
    # Preenchendo o formulário com os dados do usuário atual
    navegador.find_element_by_id("first-name").send_keys(first_name)
    navegador.find_element_by_id("last-name").send_keys(last_name)
    navegador.find_element_by_id("address").send_keys(address)
    navegador.find_element_by_id("city").send_keys(city)
    navegador.find_element_by_id("state").send_keys(state)
    navegador.find_element_by_id("zip").send_keys(zip_code)
    navegador.find_element_by_id("phone-number").send_keys(phone_number)
    navegador.find_element_by_id("email").send_keys(email)
    
    # Submeter o formulário
    navegador.find_element_by_xpath('/html/body/app-root/div[2]/app-rpa1/div/div[2]/form/input').click()
    
    # Esperar um pouco após enviar o formulário
    time.sleep(2)
