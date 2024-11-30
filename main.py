from time import sleep
import win32com.client as win32
import schedule
from selenium import webdriver
from selenium.webdriver.common.by import By



def previsao_tempo():
    #Escolhendo e mantendo navegador aberto
    opcoes = webdriver.ChromeOptions()
    opcoes.add_experimental_option("detach", True)
    navegador = webdriver.Chrome(options=opcoes)
    #Entrando no site
    navegador.get('LINK DA PREVISÃO DO TEMPO DA SUA CIDADE NO GOOGLE')


    #temperatura atual
    temp_atual = navegador.find_element(By.XPATH, '//*[@id="wob_tm"]').text
    temp_condicao = navegador.find_element(By.XPATH, '//*[@id="wob_dc"]').text
    print(f'{temp_atual}°C')
    print(temp_condicao)

    #Previsão de amanhã
    navegador.find_element(By.XPATH, '//*[@id="wob_dp"]/div[2]').click()
    temp_max_dia1 = navegador.find_element(By.XPATH, '//*[@id="wob_dp"]/div[2]/div[3]/div[1]/span[1]').text
    temp_min_dia1 = navegador.find_element(By.XPATH, '//*[@id="wob_dp"]/div[2]/div[3]/div[2]/span[1]').text
    temp_condicao_dia1 = navegador.find_element(By.XPATH, '//*[@id="wob_dc"]').text
    print(f'{temp_min_dia1}°C')
    print(f'{temp_max_dia1}°C')
    print(temp_condicao_dia1)

    #Previsão depois de amanhã
    navegador.find_element(By.XPATH, '//*[@id="wob_dp"]/div[3]').click()
    temp_max_dia2 = navegador.find_element(By.XPATH, '//*[@id="wob_dp"]/div[3]/div[3]/div[1]/span[1]').text
    temp_min_dia2 = navegador.find_element(By.XPATH, '//*[@id="wob_dp"]/div[3]/div[3]/div[2]/span[1]').text
    temp_condicao_dia2 = navegador.find_element(By.XPATH, '//*[@id="wob_dc"]').text
    print(f'{temp_min_dia2}°C')
    print(f'{temp_max_dia2}°C')
    print(temp_condicao_dia2)

    #Fechando navegado depois de 5 segundos
    sleep(5)
    navegador.quit()

    #Criando o email
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    #Informações do email
    email.To = 'EMAIL DE DESTINO'
    email.Cc = 'EMAIL DE CÓPIA'
    email.Subject = 'Previsão do tempo'
    email.Body = (f'Temperatura atual {temp_atual}°C e o tempo esta {temp_condicao}.\n'
                  f'Amanhã a temperatura mínima amanhã  será {temp_min_dia1}°C e a máxima será {temp_max_dia1}°C e o tempo vai estar {temp_condicao_dia1}.\n'
                  f'Depois de amanhã a temperatura mínima será {temp_min_dia2}°C e a máxima será {temp_max_dia2}°C e o tempo vai estar {temp_condicao_dia2}\n'
                  f'Até amanhã.')
    email.Send()

schedule.every().day.at("6:30").do(previsao_tempo)

while True:
    schedule.run_pending()
    sleep(60)
