from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
from PIL import Image
import time


servico = Service(ChromeDriverManager().install())

# Configurar o driver do Selenium (neste exemplo, usaremos o Chrome)
driver = webdriver.Chrome(service=servico)

# Carregar a página da web desejada
driver.get("https://home.portalsas.com.br/")

# Obter a altura total da página
altura_total = driver.execute_script("return document.documentElement.scrollHeight")

# Configurar a altura da janela do navegador para corresponder à altura total da página
driver.set_window_size(1366, altura_total)

# Capturar a captura de tela da página inteira
screenshot = driver.get_screenshot_as_png()

# Fechar o driver do Selenium
driver.quit()

# Salvar a captura de tela como um arquivo de imagem
with open("screenshot.png", "wb") as file:
    file.write(screenshot)
