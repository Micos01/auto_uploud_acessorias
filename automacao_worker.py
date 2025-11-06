# Imports do Python
import pandas as pd
import time
import sys 
import os 
import json # Usado para comunicação

# Imports de Terceiros
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, TimeoutException

# --- Funções de Comunicação com Rust ---

def send_to_rust(data_type, message):
    """Envia uma mensagem JSON para o stdout (que o Rust está lendo)."""
    try:
        payload = json.dumps({"type": data_type, "message": message})
        print(payload, flush=True)
    except Exception as e:
        # Se falhar, manda um erro bruto
        print(json.dumps({"type": "log", "message": f"ERRO INTERNO DO PYTHON (send_to_rust): {e}"}), flush=True)

def log(message):
    """Envia um log para o Rust."""
    send_to_rust("log", message)

def send_state(state):
    """Envia uma mudança de estado para o Rust."""
    send_to_rust("state", state)

# --- Classe de Automação (Backend) ---

class AutomationController:
    """Mantém o estado da automação (driver, wait)."""
    
    def __init__(self):
        self.driver = None
        self.wait = None
        self.stop_requested = False
        # Mapa de departamentos (do seu script original)
        self.dept_map = {
            "Administrativo": "4", "Contábil": "1", "  » Apuração Lucro Presumido": "9",
            "  » Apuração Lucro Real": "8", "  » Envio guias Lucro Presumido": "15",
            "  » Envio guias Lucro Real": "6", "Direção": "18", "Fiscal": "2",
            "  » Auditoria": "14", "  » DEFIS": "11", "  » DIMOB": "12", "  » DMED": "13",
            "  » Fiscal Gestão": "10", "Legalização": "5", "  » Alvará": "16",
            "Pessoal": "3", "TI": "17"
        }

    def get_driver_path(self):
        """Encontra o chromedriver (essencial para PyInstaller)."""
        if getattr(sys, 'frozen', False):
            application_path = sys._MEIPASS #type: ignore
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
        
        # O Rust vai procurar o chromedriver na raiz, junto com o worker
        return os.path.join(application_path, "chromedriver.exe")

    def handle_login(self, login_val, pass_val):
        """Etapa 1: Inicia o driver e faz o login."""
        try:
            log("Iniciando automação...")
            if not login_val or not pass_val:
                log("ERRO: Login ou Senha não recebidos do Rust.")
                raise Exception("Login/Senha vazios")

            chrome_options = webdriver.ChromeOptions()
            chrome_options.set_capability('acceptInsecureCerts', True)
            chrome_options.add_experimental_option("detach", True)
            driver_path = self.get_driver_path()
            
            if not os.path.exists(driver_path):
                log(f"ERRO CRÍTICO: chromedriver.exe não encontrado em {driver_path}")
                raise FileNotFoundError("O 'chromedriver.exe' não foi encontrado.")
                
            service = Service(driver_path)
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.wait = WebDriverWait(self.driver, 10)
            self.driver.get("https://app.acessorias.com/")
            
            login_field = self.driver.find_element(By.NAME, "mailAC")
            login_field.send_keys(login_val)
            pass_field = self.driver.find_element(By.NAME, "passAC")
            pass_field.send_keys(pass_val)
            
            log("\n" + "="*50)
            log("!!! AÇÃO MANUAL NECESSÁRIA !!!")
            log("Por favor, resolva o CAPTCHA na janela do navegador.")
            log("Após clicar em 'Entrar', clique no botão '2. Resolvi o CAPTCHA'")
            log("="*50 + "\n")
            
            # Envia o novo estado para o Rust
            send_state("WaitingForCaptcha")

        except Exception as e:
            log(f"ERRO na Etapa 1: {e}")
            send_state("Error")

    def handle_uploads(self, cnpj_val, dept_name, mode, csv_path, folder_path, filter_string):
        """Etapa 2: Continua a automação e faz os uploads."""
        try:
            self.stop_requested = False
            
            if not self.driver or not self.wait:
                log("ERRO CRÍTICO: Driver não foi iniciado. Tente o login novamente.")
                raise Exception("Driver não existe")

            try:
                log("Navegando para o Dashboard para garantir o início...")
                self.driver.get("https://app.acessorias.com/sysmain.php")
            except WebDriverException as e:
                log(f"ERRO CRÍTICO: Não foi possível conectar ao navegador. {e}")
                log("O navegador pode ter sido fechado. Por favor, inicie o login novamente.")
                send_state("Error")
                return 

            log("Aguardando login ser processado...")
            self.wait.until(EC.title_contains("Acessórias, automação e gestão online"))
            
            xpath_empresas = "//a[contains(., 'Empresas')]"
            link_empresas = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath_empresas)))
            link_empresas.click()
            log("Clicou no menu 'Empresas'.")

            # Pega valor do departamento
            if dept_name not in self.dept_map:
                log(f"ERRO: Departamento '{dept_name}' não é um valor válido.")
                raise Exception("Departamento inválido selecionado.")
            dept_val = self.dept_map[dept_name] 

            # Busca empresa
            self.driver.find_element(By.ID, "searchString").send_keys(cnpj_val)
            self.driver.find_element(By.ID, "btFilter").click()
            log("Aguardando resultado da busca...")
            
            # ATENÇÃO: Hardcoded para "CONNECT EMPRESAS" (do seu script original)
            # Você pode querer tornar isso dinâmico se necessário
            xpath_resultado = "//div[contains(@class, 'aImage') and contains(., 'CONNECT EMPRESAS TELECOMUNICACOES LTDA')]"
            resultado_empresa = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath_resultado)))
            resultado_empresa.click()
            log("Clicou no resultado da busca da empresa.")
            time.sleep(1) 
            
            # --- Lógica Dupla para carregar o DataFrame ---
            df = None
            
            if mode == "folder":
                log(f"Iniciando varredura da pasta: {folder_path}...")
                file_list = []
                for root, dirs, files in os.walk(folder_path):
                    for file_name in files:
                        full_path = os.path.join(root, file_name) 
                        try:
                            file_stats = os.stat(full_path)
                            file_list.append({
                                'FullName': full_path, 'Name': file_name,
                                'PSIsContainer': False, 'Length': file_stats.st_size
                            })
                        except OSError as e:
                            log(f"AVISO: Não foi possível ler o arquivo {full_path}. Ignorando. Erro: {e}")

                if not file_list:
                    log(f"ERRO: Nenhum arquivo encontrado na pasta '{folder_path}'.")
                    raise Exception("Nenhum arquivo encontrado")
                
                log(f"Varredura concluída. {len(file_list)} arquivos encontrados.")
                df = pd.DataFrame(file_list)
                
                # Bloco de Filtro
                keywords = [k.strip().lower() for k in filter_string.split(',') if k.strip()]
                
                if keywords:
                    log(f"Aplicando filtro de palavras-chave: {keywords}")
                    
                    def filter_func(row):
                        file_name_lower = str(row['Name']).lower()
                        return any(keyword in file_name_lower for keyword in keywords)
                        
                    df = df[df.apply(filter_func, axis=1)]
                    log(f"Filtro aplicado. {len(df)} arquivos restantes.")
                else:
                    log("Nenhum filtro de palavras-chave aplicado.")
            
            elif mode == "csv":
                log(f"Carregando CSV de: {csv_path}")
                try:
                    df = pd.read_csv(csv_path, encoding='utf-8-sig', engine='python')
                    log("...CSV lido com encoding 'utf-8-sig'.")
                except Exception:
                    log("...Falha no utf-8-sig. Tentando com latin1...")
                    df = pd.read_csv(csv_path, encoding='latin1', engine='python')
                    log("...CSV lido com encoding 'latin1'.")
            
            # --- Lógica de Upload ---
            if df is None or df.empty:
                log("Nenhum arquivo para upload (lista vazia ou filtro não retornou nada).")
            else:
                arquivos_para_upload = df[df['PSIsContainer'] == False].copy()
                log(f"\nEncontrados {len(arquivos_para_upload)} arquivos para fazer upload.")
                total_files = len(arquivos_para_upload)

                for index, row in arquivos_para_upload.iterrows():
                    
                    if self.stop_requested:
                        log("Processo interrompido pelo usuário.")
                        send_state("Finished")
                        break 
                    
                    file_path = row['FullName']
                    file_name = row['Name']
                    log(f"Iniciando upload para: {file_name}...")
                    
                    # Envia o progresso para o Rust
                    progress = (index + 1) / total_files
                    send_to_rust("progress", {"progress": progress, "file": file_name})
                    
                    try:
                        # --- Preenchimento dos campos ---
                        colocar_arquivo = self.wait.until(EC.presence_of_element_located((By.ID, "EmpAnexo")))
                        # Ações de JS para garantir visibilidade (do seu script original)
                        self.driver.execute_script("arguments[0].style.display = 'block';", colocar_arquivo)
                        self.driver.execute_script("arguments[0].style.visibility = 'visible';", colocar_arquivo)
                        time.sleep(1) # Pequena pausa é boa
                        colocar_arquivo.send_keys(file_path)

                        descricao_input = self.driver.find_element(By.ID, "newComAnx")
                        self.driver.execute_script("arguments[0].value = arguments[1];", descricao_input, file_name)
                        time.sleep(0.3) 

                        lista_de_selecao = self.driver.find_element(By.ID, "AnxDptID")
                        departamento_select = Select(lista_de_selecao)
                        departamento_select.select_by_value(dept_val) 

                        # --- Clique e Lógica de Espera ---
                        btn_xpath = "//button[contains(@onclick, 'addAnx')]"
                        btn_adicionar = self.wait.until(EC.element_to_be_clickable((By.XPATH, btn_xpath)))
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", btn_adicionar)
                        time.sleep(0.2) 
                        self.driver.execute_script("arguments[0].click();", btn_adicionar)
                        
                        # ETAPA 1: Esperar "Salvando anexo..."
                        log("...Upload iniciado, aguardando pop-up de carregamento...")
                        try:
                            loading_popup = WebDriverWait(self.driver, 10).until(
                                EC.visibility_of_element_located((
                                    By.XPATH, "//div[contains(@class, 'swal2-popup')]//h2[contains(text(), 'Salvando anexo')]"
                                ))
                            )
                            log("...Pop-up de carregamento detectado.")
                        except TimeoutException:
                            log("!!! ERRO: Pop-up 'Salvando anexo...' não apareceu.")
                            try:
                                # Tenta capturar um erro de validação imediato
                                error_popup = self.driver.find_element(By.CLASS_NAME, "swal2-container")
                                error_msg = error_popup.find_element(By.XPATH, ".//*[contains(@id, 'swal2-title') or contains(@id, 'swal2-content')]").text
                                log(f"!!! ERRO DE VALIDAÇÃO: {error_msg}")
                                raise Exception(f"Erro de validação da página: {error_msg}")
                            except:
                                raise Exception("Falha ao iniciar o upload (pop-up de loading não encontrado)")

                        # ETAPA 2: Esperar "Salvando anexo..." DESAPARECER
                        log("...Aguardando finalização do upload (pop-up de carregamento sumir)...")
                        WebDriverWait(self.driver, 90).until(
                            EC.invisibility_of_element(loading_popup)
                        )
                        log("...Upload finalizado. Aguardando pop-up de resultado...")
                        
                        # ETAPA 3: Esperar pop-up de RESULTADO
                        popup_container = WebDriverWait(self.driver, 10).until(
                            EC.visibility_of_element_located((By.CLASS_NAME, "swal2-container"))
                        )

                        # ETAPA 4: Verificar SUCESSO ou ERRO
                        try:
                            popup_container.find_element(By.XPATH, ".//*[contains(@id, 'swal2-title') and (contains(text(), 'realizado com sucesso') or contains(text(), 'sucesso'))]")
                            log(f"Sucesso: {file_name} enviado.")
                            # Fecha o pop-up de sucesso
                            popup_container.find_element(By.CLASS_NAME, "swal2-confirm").click()
                            self.wait.until(EC.invisibility_of_element(popup_container))
                        except Exception: 
                            error_msg = "Mensagem de erro não encontrada"
                            try: error_msg = popup_container.find_element(By.ID, "swal2-content").text
                            except:
                                try: error_msg = popup_container.find_element(By.ID, "swal2-title").text
                                except: pass
                            log(f"!!! ERRO DE UPLOAD: {error_msg}")
                            # Fecha o pop-up de erro
                            try: popup_container.find_element(By.CLASS_NAME, "swal2-confirm").click()
                            except: pass 
                            self.wait.until(EC.invisibility_of_element(popup_container))
                            raise Exception(f"Erro de validação da página: {error_msg}")

                    except Exception as e_file:
                        log(f"!!! ERRO ao enviar o arquivo {file_name}: {e_file}")
                        log("...Tentando recarregar a página e continuar...")
                        self.driver.refresh()
                        log("Página recarregada. Aguardando para continuar...")
                        time.sleep(3) 
                        continue 
            
            if not self.stop_requested:
                log("\n" + "="*50)
                log("Upload de todos os arquivos concluído!")
                log("="*50)
            
            send_state("Finished")

        except Exception as e:
            log(f"ERRO na Etapa 2: {e}")
            send_state("Error")

# --- Ponto de Entrada do Worker Python ---
def main():
    """Função principal que ouve comandos do Rust via stdin."""
    controller = AutomationController()
    
    # Envia sinal de que está pronto
    log("Worker Python iniciado e pronto para comandos.")
    
    while True:
        try:
            # Lê uma linha do stdin (enviada pelo Rust)
            line = sys.stdin.readline()
            if not line:
                # O stdin fechou, o Rust provavelmente saiu
                break 
                
            # Parseia o comando JSON
            command = json.loads(line)
            
            if command.get("action") == "login":
                data = command.get("data", {})
                controller.handle_login(data.get("login"), data.get("password"))
            
            elif command.get("action") == "continue":
                data = command.get("data", {})
                controller.handle_uploads(
                    data.get("cnpj"),
                    data.get("dept_name"),
                    data.get("mode"),
                    data.get("csv_path"),
                    data.get("folder_path"),
                    data.get("filter_keywords")
                )
            
            elif command.get("action") == "stop":
                log("Parada recebida.")
                controller.stop_requested = True
            
            else:
                log(f"Comando desconhecido: {command}")
                
        except json.JSONDecodeError:
            log(f"ERRO: Recebido JSON inválido do Rust: {line.strip()}")
        except Exception as e:
            log(f"ERRO CRÍTICO NO LOOP PRINCIPAL DO PYTHON: {e}")
            send_state("Error")

if __name__ == "__main__":
    main()