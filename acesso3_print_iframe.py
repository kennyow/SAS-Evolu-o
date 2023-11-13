from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import Keys, ActionChains
from selenium.webdriver.common.by import By
from PIL import Image
from time import sleep
import pyautogui
from docx import Document
import paragraphs
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from Screenshot import Screenshot

# -*- coding: utf-8 -*-

# Configurar o driver do Selenium (neste exemplo, usaremos o Chrome)
# Modo 'Abre e Fecha' do Chrome corrigido

options = webdriver.ChromeOptions()
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(options=options, service=Service(ChromeDriverManager().install()))

# Carregar a página da web desejada
driver.get("https://portalsas.com.br/login?redirect=https%3A%2F%2Fhome.portalsas.com.br%2F")
driver.maximize_window() # Modo Tela Cheia

# Carregar usuário/senha
pyautogui.click(x=504, y=399)
pyautogui.write("kennyo.cavalcante@colegio-evolucao.com")
pyautogui.click(x=504, y=508)
pyautogui.write("Kennyow86")

# Confirmar
driver.find_element('xpath', '//*[@id="login-form-submit"]').click()
sleep(3)

# Fechando Pop ups
WebDriverWait(driver, 30).until(
EC.presence_of_element_located(('xpath', '//*[@id="root"]/div[1]/div/div')) 
)

pyautogui.click(x=408, y=671)
pyautogui.click(x=1188, y=387)
driver.find_element('xpath', '//*[@id="root"]/nav/div[2]/button').click()

# Acessando 
driver.find_element('xpath', '//*[@id="root"]/nav/div[1]/div[1]/ul/li[6]/button').click()
driver.find_element('xpath', '//*[@id="root"]/nav/div[1]/div[1]/ul/li[6]/div/ul/li[6]/a').click()

# No Studos
sleep(12)
avali = str("A2 - QUÍMICA E BIOLOGIA - 2ª Série II Tri. 2023").upper()

# Troca de Janela
window_name = driver.window_handles[0]
driver.switch_to.window(window_name=window_name)
driver.close()
sleep(2)
driver.switch_to.window(driver.window_handles[0])
sleep(1)


# Inserindo a disciplina no box
pyautogui.click(x=369, y=444)
# Create an ActionChains object
actions = ActionChains(driver)
actions.send_keys(avali)
actions.perform()
#pyautogui.write(avali)
sleep(3)
#driver.refresh() Reload da página que só carregava os exercícios de A2
sleep(6)


# Acessando os Resultados da atividade
pyautogui.click(x=1204, y=699)
pyautogui.click(x=1033, y=647)
sleep(7)



# Find all <td> elements with any class
sleep(5)
lista_percentagem = []
td_texts = []
'''xpathe = '/html/body/div[2]/div[1]/div[2]/div/div[3]/div/div[3]/div[2]/div[4]/div/div/div/div/div/div'
td_elements = driver.find_element(By.XPATH, xpathe)'''
xpathe = '/html/body/div[2]/div[1]/div[2]/div/div[3]/div/div[3]/div[2]/div[1]/div/div/div/ul/li[3]/a'
td_elements = driver.find_element(By.XPATH, xpathe)
actions = ActionChains(driver)
actions.move_to_element(td_elements).perform()
td_elements.click()
sleep(3)

#   Questão 1
xpathe2 = '//*[@id="questionReport"]/div/div/div/div/div/div/div/div/div[1]/table/tbody/tr[1]/td[1]'
td_elements = driver.find_element(By.XPATH, xpathe2)
actions = ActionChains(driver)
actions.move_to_element(td_elements).perform()
td_elements.click()
sleep(3)

# Iframe Pop Up

'''driver.find_element(By.CLASS_NAME, "ReactModal__Content ReactModal__Content--after-open").click()
sleep(1)'''




# frame = WebDriverWait(driver, 10).until(EC.frame_to_be_available_and_switch_to_it(By.ID, xpathe3))
'''actions = ActionChains(driver)
actions.move_to_element(td_elements).perform()'''

ob = Screenshot.Screenshot()
img_url = ob.full_screenshot(driver, save_path=r'.', image_name='myimage.png', is_load_at_runtime=True,
                                          load_wait_time=4)
driver.find_element(By.XPATH, '/html/body/div[5]/div/div/div/div[1]/div[2]/div/div/span/svg').click()
sleep(2)


'''for numero in range(1,31):
    #xpathe = '/html/body/div[2]/div[1]/div[2]/div/div[3]/div/div[3]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div/div[1]/div[2]/div/div['+ str(numero)+']/button/svg/path[2]'
    
    num = int(td_elements.text.strip()[:-1])
    lista_percentagem.append(num)
    
    for element in td_elements:
        
    /html/body/div[2]/div[1]/div[2]/div/div[3]/div/div[3]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div/div[1]/div[2]/div/div[1]/button/svg/path[2]
    /html/body/div[2]/div[1]/div[2]/div/div[3]/div/div[3]/div[2]/div[2]/div/div/div/div/div/div/div[2]/div/div[1]/div[2]/div/div[2]/button/svg/path[2]









        #print(element.text.strip())
        #print(num)
        if num < 60:
            //*[@id="questionReport"]/div/div/div/div/div/div/div/div/div[1]/table/tbody/tr[1]/td[1]
            //*[@id="questionReport"]/div/div/div/div/div/div/div/div/div[1]/table/tbody/tr[2]/td[1]
            actions = ActionChains(driver)
            actions.move_to_element(td_elements).perform()
            #amiguri.click()
          
            x_position = td_elements.location['x']
            y_position = td_elements.location['y']
            pyautogui.click(x_position, y_position)
            sleep(10)
            print(f"Width: {x_position}px")
            print(f"Height: {y_position}px")'''
'''td_elements = driver.find_elements(By.XPATH, '//*[@id="questionReport"]/div/div/div/div/div/div/div/div/div[1]/table/tbody/tr[2]/td[5]'''     


# Extract the text from each element and put them into a list
#td_texts = [element.text for element in td_elements]
#print(lista_percentagem)
# Print the list of extracted texts
#print(td_texts)

'''
# Print página inteira
ob = Screenshot.Screenshot()
img_url = ob.full_screenshot(driver, save_path=r'.', image_name='myimage.png', is_load_at_runtime=True,
                                          load_wait_time=4)'''