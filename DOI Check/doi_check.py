from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
import time
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
import os

# Carrega variáveis de ambiente do arquivo .env
load_dotenv()

# Configura driver
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# Abre a página
driver.get("https://rilavras-mg.sistemaasgard.com.br/app/v1/#/tarefas")
time.sleep(3)  # espera carregar

# Preenche o usuário
input_usuario = driver.find_element(By.ID, "username")
input_usuario.send_keys("gabriel.zanda")

# Preenche a senha
input_senha = driver.find_element(By.ID, "password")
input_senha.send_keys(os.getenv("PASSWORD"))

input_senha.send_keys(Keys.RETURN)
time.sleep(5)  # espera carregar

# Abrir modal de busca com Ctrl+Alt+F
body = driver.find_element(By.TAG_NAME, "body")
body.send_keys(Keys.CONTROL, Keys.ALT, "f")

# Agora localiza o input 
wait = WebDriverWait(driver, 10)
input_busca = wait.until(EC.visibility_of_element_located(
    (By.CSS_SELECTOR, "input[placeholder='Faça sua busca'].form-control.form-control-xs")
))

# Insere a matrícula
matricula = "79.725"
matricula_extenso = "Matrícula 79.725"
input_busca.send_keys(matricula)
input_busca.send_keys(Keys.RETURN)  # Pressiona Enter para buscar

# Aguarda carregar os resultados
wait = WebDriverWait(driver, 10)

# Espera o dropdown aparecer (a div que você indicou)
dropdown = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "div.search-results.dropdown-menu")))

# Localiza o link desejado dentro do dropdown (exemplo pegando pelo texto parcial)
item = dropdown.find_element(By.XPATH, ".//a[contains(., 'Matrícula 79.725')]")

# Clica no item para abrir a matrícula
item.click()

# Aguarda carregar os resultados
wait = WebDriverWait(driver, 10)

# Vai para "Atos Registrados"
atos_registrados = wait.until(
    EC.element_to_be_clickable((By.XPATH, "//a[@title='Atos Registrados']"))
)

#Click
atos_registrados.click()

# Clicar no ato específico
elemento = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((
        By.XPATH,
        "//strong[contains(normalize-space(), 'M.0')]"
    ))
)
elemento.click()

input("Pressione Enter para fechar o navegador...")
driver.quit()