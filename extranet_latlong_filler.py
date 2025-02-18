import pandas as pd
import selenium
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from geopy.geocoders import Nominatim
from geopy.exc import GeocoderTimedOut, GeopyError

# Ler o arquivo Excel
ids_produtos = pd.read_excel('ProdutosExtranet.xlsx')
coluna = 'ID'

lista_produtos = ids_produtos[coluna].tolist()
print(lista_produtos)

url_Extranet = 'https://extranet.lopesrio.com.br/'
login_Extranet = "141.984.977-88"
senha_Extranet = "wi27vtczZzZ!"

#Navegador
navegador = webdriver.Chrome()
navegador.get(url_Extranet)
navegador.maximize_window()
time.sleep(5)
navegador.find_element(By.XPATH, '/html/body/form/div[3]/div/section/article/fieldset/label[1]/input').send_keys(login_Extranet)
navegador.find_element(By.XPATH , '/html/body/form/div[3]/div/section/article/fieldset/label[2]/input').send_keys(senha_Extranet)
navegador.find_element(By.XPATH, '/html/body/form/div[3]/div/section/article/fieldset/label[3]/a').click()
time.sleep(5)

geolocator = Nominatim(user_agent='geoapi_exercises', timeout=10)

# Iterar sobre todos os IDs de produtos e acessar o sistema
for produto_id in lista_produtos:
    url_empreendimentos = f'https://extranet.lopesrio.com.br/ERP/EMP_Cadastro.aspx?emp={produto_id}'

    
    navegador.get(url_empreendimentos)
    time.sleep(5)

    try:
        # Montar endereço completo
        endereço = navegador.find_element(By.ID, 'ctl00_ContentPlaceHolder1_tEndereco').get_attribute('value')
        numero = navegador.find_element(By.ID, 'ctl00_ContentPlaceHolder1_tNumero').get_attribute('value')
        bairro = navegador.find_element(By.ID, 'ctl00_ContentPlaceHolder1_tBairro').get_attribute('value')
        cep = navegador.find_element(By.ID, 'ctl00_ContentPlaceHolder1_tCEP').get_attribute('value')

        # Verificar campos em branco
        if not endereço or not numero or not bairro or not cep:
            print(f"Dados incompletos para o ID {produto_id}: Endereço({endereço}), Número({numero}), Bairro({bairro}), CEP({cep}).")
            continue 

        endereço_final = f'{endereço}, {numero} - {bairro} - {cep}'
        print(f"Processando: {endereço_final}")

        # Geocode
        try:
            location = geolocator.geocode(endereço_final)

            if location:
                lat = location.latitude
                long = location.longitude
                print(f"Latitude: {lat}, Longitude: {long}")
            else:
                print(f"Endereço não encontrado: {endereço_final}.")
                continue
        except (GeocoderTimedOut, GeopyError) as e:
            print(f"Erro de geocodificação para {endereço_final}: {str(e)}")
            continue

        # Substituir as latitudes e longitudes anteriores pelas novas
        campo_latitude = navegador.find_element(By.ID, 'ctl00_ContentPlaceHolder1_tLatitude')
        campo_longitude = navegador.find_element(By.ID, 'ctl00_ContentPlaceHolder1_tLongitude')
        campo_latitude.send_keys(Keys.CONTROL, 'a')
        campo_latitude.send_keys(Keys.BACKSPACE)
        campo_longitude.send_keys(Keys.CONTROL, 'a')
        campo_longitude.send_keys(Keys.BACKSPACE)
        campo_latitude.send_keys(str(lat))
        campo_longitude.send_keys(str(long))
        navegador.find_element(By.ID, 'ctl00_ContentPlaceHolder1_btn').click()
        time.sleep(2)  

    except selenium.common.exceptions.NoSuchElementException as e:
        print(f"Erro ao encontrar elementos necessários: {str(e)}")
        continue

# Fechar o navegador após o processo
navegador.quit()

print("Script completado com sucesso!")