
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd


def main():
     
     chrome_options = webdriver.ChromeOptions()
     chrome_options.set_capability('acceptInsecureCerts', True)
     # Garante que a janela do navegador não feche ao final do script (útil para depuração)
     chrome_options.add_experimental_option("detach", True)

     service = Service(ChromeDriverManager().install())
     driver = webdriver.Chrome(service=service, options=chrome_options)

     driver.get("https://app.acessorias.com/")
     wait = WebDriverWait(driver, 10)

     # Preenche o login e senha
     login = driver.find_element(By.NAME, "mailAC")
     login.send_keys("ti.2@merca.com.br")

     senha = driver.find_element(By.NAME, "passAC")
     senha.send_keys("Merca@2025_acessorias")

     # Pausa para resolver o CAPTCHA manualmente
     print("\n" + "="*50)
     print("!!! AÇÃO MANUAL NECESSÁRIA !!!")
     print("Por favor, resolva o CAPTCHA na janela do navegador.")
     input("Após resolver e clicar em 'Entrar', pressione Enter aqui para continuar...")
     print("="*50 + "\n")

     print("Continuando a automação após o login...")


     # Agora você pode adicionar o código para interagir com a página já logada.
     # Exemplo: esperar pelo título do dashboard para confirmar o login
     wait.until(EC.title_contains("Acessórias, automação e gestão online de prazos e processos para sua empresa contábil"))
     xpath_empresas = "//a[contains(., 'Empresas')]"
     link_empresas = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_empresas)))
     link_empresas.click()
     print("Clicou no menu 'Empresas'.")


     pesquisa_empresa = driver.find_element(By.NAME, "search")         
     pesquisa_empresa.send_keys("28.360.182/0001-20")

     btn_filtrar = driver.find_element(By.ID, "btFilter")

     btn_filtrar.click()

     # Espera o resultado da busca aparecer e clica nele
     print("Aguardando resultado da busca...")
     xpath_resultado = "//div[contains(@class, 'aImage') and contains(., 'CONNECT EMPRESAS TELECOMUNICACOES LTDA')]"
     resultado_empresa = wait.until(EC.element_to_be_clickable((By.XPATH, xpath_resultado)))
     resultado_empresa.click()
     print("Clicou no resultado da busca da empresa.")
     ativar_area = driver.find_element(By.XPATH,"//label[contains(.,'Arquivos anexos')]")
     ativar_area.click()

     wait = WebDriverWait(driver, 10)
     colocar_arquivo = wait.until(EC.presence_of_element_located((By.ID,"EmpAnexo")))

     colocar_arquivo.send_keys()
     lista_de_selecao = driver.find_element(By.ID, "AnxDptID")
     departamento_select = Select(lista_de_selecao)
     departamento_select.select_by_value("1")





if __name__ == "__main__":
    main()
