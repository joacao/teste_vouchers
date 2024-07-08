from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import time
import pandas as pd
import datetime
import os
from dotenv import load_dotenv

load_dotenv(override = True)

email = os.getenv('email')
senha = os.getenv('senha')

data_atual = datetime.date.today()
data_formatada = data_atual.strftime('%d-%m-%Y')

cliente = input('Qual o cliente? ')
iniciativa = input('Qual iniciativa você quer testar os vouchers [L1],[L2] ou [TSA]: ')

servico = Service()
options = webdriver.FirefoxOptions()
navegador = webdriver.Firefox(service = servico, options = options)


PLANILHA_VOUCHERS = pd.read_excel(r'~/automacao/teste_vouchers/Vouchers para testar.xlsx',engine='openpyxl')


#1 - Acessar o link 
navegador.get("https://accounts.google.com/o/oauth2/auth/oauthchooseaccount?client_id=345223524425-5kfdrbssk5q4sgn2k5r9m798h6e19cue.apps.googleusercontent.com&redirect_uri=https%3A%2F%2Fwww.webassessor.com%2Fform%2FOAuth2Callback.do%3Fmethod%3Dprocess&response_type=code&scope=https%3A%2F%2Fwww.googleapis.com%2Fauth%2Fuserinfo.email%20https%3A%2F%2Fwww.googleapis.com%2Fauth%2Fuserinfo.profile&state=google-GOOGLEPTBR&service=lso&o2v=1&flowName=GeneralOAuthFlow")

#2 - Iniciar a seção com a conta Google
navegador.find_element('xpath','//*[@id="identifierId"]').send_keys(email)
time.sleep(5)
navegador.find_element('xpath','//*[@id="identifierNext"]/div/button/span').click()
time.sleep(5)
navegador.find_element('xpath','//*[@id="password"]/div[1]/div/div[1]/input').send_keys(senha)
time.sleep(5)
navegador.find_element('xpath','//*[@id="passwordNext"]/div/button/span').click()
time.sleep(10)

#3 - Ir até a "inscrição para o exame"
navegador.find_element('xpath','//*[@id="7"]/a').click()
time.sleep(5)
vouchers = []
preco_vouchers = []
status_voucher = []

#Selecionar a prova
if iniciativa == 'L1':
    for voucher in PLANILHA_VOUCHERS['Voucher Code']:    
        try:
            navegador.find_element('xpath','//*[@id="cat_25089"]/div[2]/div[6]/a').click()
            time.sleep(5)
            navegador.find_element('xpath','//*[@id="shopping"]/div[2]/div/div/form/table/tbody/tr[4]/td[2]/input').send_keys(voucher)
            time.sleep(5)
            navegador.find_element('xpath','//*[@id="btnSubmit"]').click()
            time.sleep(5)
            dados_preco = navegador.find_element('xpath','//*[@id="spTotalPrice"]').text
            preco_vouchers.append(dados_preco)
            vouchers.append(voucher)
            navegador.find_element('xpath','//*[@id="shopping"]/div[2]/div/div/div/button').click()
            time.sleep(5)
            navegador.get('https://www.webassessor.com/wa.do?page=enterCatalog&tabs=7')
            time.sleep(5)
        except Exception as e:
            print(e,voucher)

elif iniciativa == 'L2':
    for voucher in PLANILHA_VOUCHERS['Voucher Code']:
        try:
            navegador.find_element('xpath','//*[@id="cat_25089"]/div[3]/div[6]/a').click()
            time.sleep(5)
            navegador.find_element('xpath','//*[@id="shopping"]/div[2]/div/div/form/table/tbody/tr[4]/td[2]/input').send_keys(voucher)
            time.sleep(5)
            navegador.find_element('xpath','//*[@id="btnSubmit"]').click()
            time.sleep(5)
            dados_preco = navegador.find_element('xpath','//*[@id="spTotalPrice"]').text
            preco_vouchers.append(dados_preco)
            vouchers.append(voucher)
            navegador.find_element('xpath','//*[@id="shopping"]/div[2]/div/div/div/button').click()
            time.sleep(5)
            navegador.get('https://www.webassessor.com/wa.do?page=enterCatalog&tabs=7')
            time.sleep(5)
        except Exception as e:
            print(e,voucher)

elif iniciativa == 'TSA':
    for voucher in PLANILHA_VOUCHERS['Voucher Code']:
        try:
            navegador.find_element('xpath','//*[@id="cat_25089"]/div[1]/div[6]/a').click()
            time.sleep(5)
            navegador.find_element('xpath','//*[@id="shopping"]/div[2]/div/div/form/table/tbody/tr[4]/td[2]/input').send_keys(voucher)
            time.sleep(5)
            navegador.find_element('xpath','//*[@id="btnSubmit"]').click()
            time.sleep(5)
            dados_preco = navegador.find_element('xpath','//*[@id="spTotalPrice"]').text
            preco_vouchers.append(dados_preco)
            vouchers.append(voucher)
            navegador.find_element('xpath','//*[@id="shopping"]/div[2]/div/div/div/button').click()
            time.sleep(5)
            navegador.get('https://www.webassessor.com/wa.do?page=enterCatalog&tabs=7')
            time.sleep(5)
        except Exception as e:
                print(e,voucher)


for valor in preco_vouchers:
    if valor == 'USD 0,00':
        valor = 'Utilizável'
        status_voucher.append(valor)
    else:
        valor = 'Usado'
        status_voucher.append(valor)


para_planilha = {
    'Voucher':vouchers,
    'Status':status_voucher,
}

df = pd.DataFrame(para_planilha)
df.to_excel(f'{cliente} - Vouchers testados {iniciativa} - {data_formatada} .xlsx', index= False)