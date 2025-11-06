# Imports do Python
import tkinter as tk
from tkinter import scrolledtext, filedialog
import threading
import pandas as pd
import time
import sys 
import os  

# Imports de Terceiros
import ttkbootstrap as ttk 
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, TimeoutException

class AutomationApp:
    
    def __init__(self, root):
        self.root = root
        self.root.title("Automação de Upload Acessórias")
        
        self.driver = None
        self.wait = None 
        self.stop_requested = threading.Event()
        
        self.frame = ttk.Frame(root, padding="10")
        self.frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # --- Linha 0: Login ---
        ttk.Label(self.frame, text="Login:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.login_entry = ttk.Entry(self.frame, width=40)
        self.login_entry.grid(row=0, column=1, columnspan=2, sticky=tk.W)
        self.setup_placeholder(self.login_entry, "seu.email@dominio.com")
        
        # --- Linha 1: Senha ---
        ttk.Label(self.frame, text="Senha:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.pass_entry = ttk.Entry(self.frame, width=40)
        self.pass_entry.grid(row=1, column=1, columnspan=2, sticky=tk.W)
        self.setup_placeholder(self.pass_entry, "Sua Senha", is_password=True)

        # --- Linha 2: CNPJ ---
        ttk.Label(self.frame, text="CNPJ:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.cnpj_entry = ttk.Entry(self.frame, width=40)
        # Deixado como exemplo, conforme solicitado
        self.cnpj_entry.insert(0, "28.360.182/0001-20") 
        self.cnpj_entry.grid(row=2, column=1, columnspan=2, sticky=tk.W)

        # -----------------------------------------------------------------
        # MUDANÇA: Linha 3: Reorganizado "Modo de Entrada"
        # -----------------------------------------------------------------
        mode_frame = ttk.Frame(self.frame)
        mode_frame.grid(row=3, column=0, columnspan=3, sticky=tk.W, pady=(10, 0))
        
        ttk.Label(mode_frame, text="Modo de Entrada:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.input_mode = tk.StringVar(value="csv") # Padrão é CSV
        
        self.csv_radio = ttk.Radiobutton(mode_frame, text="Usar Arquivo CSV", variable=self.input_mode, value="csv", command=self.toggle_input_mode)
        self.csv_radio.grid(row=0, column=1, sticky=tk.W, padx=10)
        
        self.folder_radio = ttk.Radiobutton(mode_frame, text="Escanear Pasta", variable=self.input_mode, value="folder", command=self.toggle_input_mode)
        self.folder_radio.grid(row=0, column=2, sticky=tk.W, padx=10)

        # -----------------------------------------------------------------
        # MUDANÇA: Linha 4: Frames Dinâmicos agora são LabelFrames
        # -----------------------------------------------------------------
        
        # --- Frame para ESCANEAR PASTA (inicialmente oculto) ---
        self.folder_frame = ttk.Labelframe(self.frame, text="Opções: Escanear Pasta", padding=10)
        # Colocado na Linha 4
        self.folder_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), padx=5, pady=5) 
        
        ttk.Label(self.folder_frame, text="Pasta:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.folder_path_entry = ttk.Entry(self.folder_frame, width=40)
        self.folder_path_entry.grid(row=0, column=1, sticky=tk.W, padx=5)
        self.browse_folder_btn = ttk.Button(self.folder_frame, text="Procurar...", command=self.browse_folder, style="outline")
        self.browse_folder_btn.grid(row=0, column=2, sticky=tk.W, padx=5)
        
        ttk.Label(self.folder_frame, text="Filtro (palavras-chave):").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.filter_placeholder = "ex: relatorio, nota, 2024 (separar por vírgula)"
        self.filter_entry = ttk.Entry(self.folder_frame, width=40)
        self.filter_entry.insert(0, self.filter_placeholder)
        self.filter_entry.grid(row=1, column=1, sticky=tk.W, padx=5)
        self.setup_placeholder(self.filter_entry, self.filter_placeholder) # Aplica placeholder

        # --- Frame para USAR CSV (inicialmente visível) ---
        self.csv_frame = ttk.Labelframe(self.frame, text="Opções: Usar Arquivo CSV", padding=10)
        # Colocado na Mesma Linha 4
        self.csv_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), padx=5, pady=5) 

        ttk.Label(self.csv_frame, text="Arquivo CSV:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.csv_path_entry = ttk.Entry(self.csv_frame, width=40)
        self.csv_path_entry.insert(0, "FolderInventory.csv")
        self.csv_path_entry.grid(row=0, column=1, sticky=tk.W, padx=5)
        self.browse_csv_btn = ttk.Button(self.csv_frame, text="Procurar...", command=self.browse_csv, style="outline")
        self.browse_csv_btn.grid(row=0, column=2, sticky=tk.W, padx=5)
        
        # -----------------------------------------------------------------
        # Linha 5: Departamento (linha atualizada)
        # -----------------------------------------------------------------
        ttk.Label(self.frame, text="Departamento:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.dept_map = {
            "Administrativo": "4", "Contábil": "1", "  » Apuração Lucro Presumido": "9",
            "  » Apuração Lucro Real": "8", "  » Envio guias Lucro Presumido": "15",
            "  » Envio guias Lucro Real": "6", "Direção": "18", "Fiscal": "2",
            "  » Auditoria": "14", "  » DEFIS": "11", "  » DIMOB": "12", "  » DMED": "13",
            "  » Fiscal Gestão": "10", "Legalização": "5", "  » Alvará": "16",
            "Pessoal": "3", "TI": "17"
        }
        self.dept_combo = ttk.Combobox(self.frame, width=38, values=list(self.dept_map.keys()))
        self.dept_combo.set("Legalização") # Padrão (Valor "5")
        self.dept_combo.grid(row=5, column=1, columnspan=2, sticky=tk.W)

        # --- Linhas 6 e 7: Botões (linhas atualizadas) ---
        self.start_button = ttk.Button(self.frame, text="1. Iniciar Login", command=self.start_automation, style="primary")
        self.start_button.grid(row=6, column=0, pady=10, sticky=tk.W, padx=5)
        
        self.captcha_button = ttk.Button(self.frame, text="2. Resolvi o CAPTCHA", command=self.continue_automation_threaded, state="disabled", style="success")
        self.captcha_button.grid(row=6, column=1, pady=10, sticky=tk.W, padx=5)
        
        self.restart_button = ttk.Button(self.frame, text="Reiniciar Uploads", command=self.continue_automation_threaded, state="disabled", style="info")
        self.restart_button.grid(row=7, column=0, pady=10, sticky=tk.W, padx=5)

        self.stop_button = ttk.Button(self.frame, text="Parar Automação", command=self.request_stop, state="disabled", style="danger")
        self.stop_button.grid(row=7, column=1, pady=10, sticky=tk.W, padx=5)
        
        # --- Linhas 8 e 9: Log (linhas atualizadas) ---
        ttk.Label(self.frame, text="Log de Atividades:").grid(row=8, column=0, sticky=tk.W, pady=5)
        self.log_widget = scrolledtext.ScrolledText(self.frame, height=15, width=70, wrap=tk.WORD)
        self.log_widget.grid(row=9, column=0, columnspan=3, pady=5)
        
        self.toggle_input_mode()

    # -----------------------------------------------------------------
    # NOVAS FUNÇÕES: Lógica de Placeholder
    # -----------------------------------------------------------------
    def on_focus_in(self, event):
        """Chamado quando o widget de Entry ganha foco."""
        widget = event.widget
        if widget.get() == widget.placeholder:
            widget.delete(0, tk.END)
            widget.config(style="TEntry") # Reseta para o estilo padrão (preto)
            if widget.is_password:
                widget.config(show="*")
    
    def on_focus_out(self, event):
        """Chamado quando o widget de Entry perde foco."""
        widget = event.widget
        if not widget.get():
            widget.insert(0, widget.placeholder)
            widget.config(style="secondary.TEntry") # Estilo 'secundário' (cinza)
            if widget.is_password:
                widget.config(show="") # Mostra o placeholder "Sua Senha"

    def setup_placeholder(self, widget, placeholder, is_password=False):
        """Configura um widget de Entry para ter um placeholder."""
        widget.placeholder = placeholder
        widget.is_password = is_password
        
        widget.config(style="secondary.TEntry") # Começa com estilo cinza
        widget.insert(0, placeholder)
        if is_password:
            widget.config(show="") # Garante que o placeholder seja visível
            
        widget.bind("<FocusIn>", self.on_focus_in)
        widget.bind("<FocusOut>", self.on_focus_out)
    # -----------------------------------------------------------------

    def toggle_input_mode(self):
        """Mostra/oculta os frames de CSV ou Pasta com base no Rádio Button."""
        mode = self.input_mode.get()
        if mode == "folder":
            self.folder_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), padx=5, pady=5)
            self.csv_frame.grid_remove() 
        elif mode == "csv":
            self.folder_frame.grid_remove() 
            self.csv_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), padx=5, pady=5)

    def get_driver_path(self):
        """Retorna o caminho para o chromedriver, seja em modo dev ou .exe"""
        if getattr(sys, 'frozen', False):
            application_path = sys._MEIPASS #type: ignore
        else:
            application_path = os.path.dirname(os.path.abspath(__file__))
        
        return os.path.join(application_path, "chromedriver.exe")

    def browse_folder(self):
        foldername = filedialog.askdirectory(title="Selecione a pasta raiz dos arquivos")
        if foldername:
            self.folder_path_entry.delete(0, tk.END)
            self.folder_path_entry.insert(0, foldername)
            
    def browse_csv(self):
        filename = filedialog.askopenfilename(
            title="Selecione o arquivo CSV",
            filetypes=(("CSV Files", "*.csv"), ("All Files", "*.*"))
        )
        if filename:
            self.csv_path_entry.delete(0, tk.END)
            self.csv_path_entry.insert(0, filename)

    def log(self, message):
        print(message) 
        self.log_widget.insert(tk.END, message + "\n")
        self.log_widget.see(tk.END)
        self.root.update_idletasks()

    def request_stop(self):
        self.stop_requested.set()
        self.log("!!! PARADA SOLICITADA !!! Aguardando o término do arquivo atual...")
        self.stop_button.config(state="disabled")

    def start_automation(self):
        self.log("Iniciando automação...")
        
        # -----------------------------------------------------------------
        # MUDANÇA: Validar campos de placeholder
        # -----------------------------------------------------------------
        try:
            login_val = self.login_entry.get()
            pass_val = self.pass_entry.get()

            if login_val == self.login_entry.placeholder:
                self.log("ERRO: Por favor, insira seu email.")
                raise Exception("Login não preenchido")
            
            if pass_val == self.pass_entry.placeholder:
                self.log("ERRO: Por favor, insira sua senha.")
                raise Exception("Senha não preenchida")
            # -----------------------------------------------------------------

            self.start_button.config(state="disabled")
            self.restart_button.config(state="disabled") 
            
            chrome_options = webdriver.ChromeOptions()
            chrome_options.set_capability('acceptInsecureCerts', True)
            chrome_options.add_experimental_option("detach", True)
            driver_path = self.get_driver_path()
            if not os.path.exists(driver_path):
                self.log(f"ERRO CRÍTICO: chromedriver.exe não encontrado em {driver_path}")
                raise FileNotFoundError("O 'chromedriver.exe' não foi encontrado.")
            service = Service(driver_path)
            self.driver = webdriver.Chrome(service=service, options=chrome_options)
            self.wait = WebDriverWait(self.driver, 10)
            self.driver.get("https://app.acessorias.com/")
            
            login = self.driver.find_element(By.NAME, "mailAC")
            login.send_keys(login_val)
            senha = self.driver.find_element(By.NAME, "passAC")
            senha.send_keys(pass_val)
            
            self.log("\n" + "="*50)
            self.log("!!! AÇÃO MANUAL NECESSÁRIA !!!")
            self.log("Por favor, resolva o CAPTCHA na janela do navegador.")
            self.log("Após clicar em 'Entrar', clique no botão '2. Resolvi o CAPTCHA'")
            self.log("="*50 + "\n")
            self.captcha_button.config(state="normal")
            
        except Exception as e:
            self.log(f"ERRO na Etapa 1: {e}")
            self.start_button.config(state="normal")


    def continue_automation_threaded(self):
        self.log("Continuando automação (em background)...")
        self.stop_requested.clear()
        self.start_button.config(state="disabled")
        self.captcha_button.config(state="disabled")
        self.restart_button.config(state="disabled")
        self.stop_button.config(state="normal") 
        thread = threading.Thread(target=self.continue_automation, daemon=True)
        thread.start()

    def continue_automation(self):
        try:
            try:
                self.log("Navegando para o Dashboard para garantir o início...")
                self.driver.get("https://app.acessorias.com/sysmain.php")
            except WebDriverException as e:
                self.log(f"ERRO CRÍTICO: Não foi possível conectar ao navegador. {e}")
                self.log("O navegador pode ter sido fechado. Por favor, inicie o login novamente.")
                return 

            self.log("Aguardando login ser processado...")
            self.wait.until(EC.title_contains("Acessórias, automação e gestão online de prazos e processos para sua empresa contábil"))
            
            xpath_empresas = "//a[contains(., 'Empresas')]"
            link_empresas = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath_empresas)))
            link_empresas.click()
            self.log("Clicou no menu 'Empresas'.")

            # --- Pega valores da GUI ---
            cnpj_val = self.cnpj_entry.get()
            dept_name = self.dept_combo.get()
            if dept_name not in self.dept_map:
                self.log(f"ERRO: Departamento '{dept_name}' não é um valor válido.")
                raise Exception("Departamento inválido selecionado.")
            dept_val = self.dept_map[dept_name] 

            pesquisa_empresa = self.driver.find_element(By.ID, "searchString")     
            pesquisa_empresa.send_keys(cnpj_val)
            btn_filtrar = self.driver.find_element(By.ID, "btFilter")
            btn_filtrar.click()

            self.log("Aguardando resultado da busca...")
            xpath_resultado = "//div[contains(@class, 'aImage') and contains(., 'CONNECT EMPRESAS TELECOMUNICACOES LTDA')]"
            resultado_empresa = self.wait.until(EC.element_to_be_clickable((By.XPATH, xpath_resultado)))
            resultado_empresa.click()
            self.log("Clicou no resultado da busca da empresa.")
            time.sleep(1) 
            
            # -----------------------------------------------------------------
            # Lógica Dupla para carregar o DataFrame
            # -----------------------------------------------------------------
            df = None
            mode = self.input_mode.get()
            
            if mode == "folder":
                folder_path = self.folder_path_entry.get()
                self.log(f"Iniciando varredura da pasta: {folder_path}...")
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
                            self.log(f"AVISO: Não foi possível ler o arquivo {full_path}. Ignorando. Erro: {e}")

                if not file_list:
                    self.log(f"ERRO: Nenhum arquivo encontrado na pasta '{folder_path}'.")
                    raise Exception("Nenhum arquivo encontrado na pasta de origem.")
                
                self.log(f"Varredura concluída. {len(file_list)} arquivos encontrados.")
                
                # Bloco de Filtro de Palavras-chave
                filter_string = self.filter_entry.get()
                keywords = []
                
                if filter_string and filter_string != self.filter_placeholder:
                    keywords = [k.strip().lower() for k in filter_string.split(',') if k.strip()]
                
                if keywords:
                    self.log(f"Aplicando filtro de palavras-chave: {keywords}")
                    filtered_file_list = []
                    for file_data in file_list:
                        file_name_lower = file_data['Name'].lower()
                        if any(keyword in file_name_lower for keyword in keywords):
                            filtered_file_list.append(file_data)
                    
                    self.log(f"Filtro aplicado. {len(filtered_file_list)} arquivos restantes.")
                    df = pd.DataFrame(filtered_file_list)
                else:
                    self.log("Nenhum filtro de palavras-chave aplicado.")
                    df = pd.DataFrame(file_list)
            
            elif mode == "csv":
                csv_path = self.csv_path_entry.get()
                self.log(f"Carregando CSV de: {csv_path}")
                try:
                    df = pd.read_csv(csv_path, encoding='utf-8-sig', engine='python')
                    self.log("...CSV lido com encoding 'utf-8-sig' (Padrão PowerShell).")
                except Exception as e:
                    self.log(f"Erro ao ler CSV com utf-8-sig: {e}. Tentando com latin1...")
                    df = pd.read_csv(csv_path, encoding='latin1', engine='python')
                    self.log("...CSV lido com encoding 'latin1'.")
            
            # --- O resto do script opera no 'df' gerado ---
            if df is None or df.empty:
                self.log("Nenhum arquivo para upload (lista vazia ou filtro não retornou nada).")
                pass 
            else:
                arquivos_para_upload = df[df['PSIsContainer'] == False].copy() #type: ignore
                self.log(f"\nEncontrados {len(arquivos_para_upload)} arquivos para fazer upload.")

                for index, row in arquivos_para_upload.iterrows():
                    
                    if self.stop_requested.is_set():
                        self.log("Processo interrompido pelo usuário.")
                        break 
                    
                    file_path = row['FullName']
                    file_name = row['Name']
                    self.log(f"Iniciando upload para: {file_name}...")
                    
                    try:
                        # --- Preenchimento dos campos ---
                        colocar_arquivo = self.wait.until(EC.presence_of_element_located((By.ID, "EmpAnexo")))
                        self.driver.execute_script("arguments[0].style.display = 'block';", colocar_arquivo)
                        self.driver.execute_script("arguments[0].style.visibility = 'visible';", colocar_arquivo)
                        time.sleep(1)
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
                        self.log("...Upload iniciado, aguardando pop-up de carregamento...")
                        try:
                            loading_popup = WebDriverWait(self.driver, 10).until(
                                EC.visibility_of_element_located((
                                    By.XPATH, "//div[contains(@class, 'swal2-popup')]//h2[contains(text(), 'Salvando anexo')]"
                                ))
                            )
                            self.log("...Pop-up de carregamento detectado.")
                        except TimeoutException:
                            self.log("!!! ERRO: Pop-up 'Salvando anexo...' não apareceu.")
                            try:
                                error_popup = self.driver.find_element(By.CLASS_NAME, "swal2-container")
                                error_msg = error_popup.find_element(By.XPATH, ".//*[contains(@id, 'swal2-title') or contains(@id, 'swal2-content')]").text
                                self.log(f"!!! ERRO DE VALIDAÇÃO: {error_msg}")
                                raise Exception(f"Erro de validação da página: {error_msg}")
                            except:
                                raise Exception("Falha ao iniciar o upload (pop-up de loading não encontrado)")

                        # ETAPA 2: Esperar "Salvando anexo..." DESAPARECER
                        self.log("...Aguardando finalização do upload (pop-up de carregamento sumir)...")
                        WebDriverWait(self.driver, 90).until(
                            EC.invisibility_of_element(loading_popup)
                        )
                        self.log("...Upload finalizado. Aguardando pop-up de resultado...")
                        
                        # ETAPA 3: Esperar pop-up de RESULTADO
                        popup_container = WebDriverWait(self.driver, 10).until(
                            EC.visibility_of_element_located((By.CLASS_NAME, "swal2-container"))
                        )

                        # ETAPA 4: Verificar SUCESSO ou ERRO
                        try:
                            popup_container.find_element(By.XPATH, ".//*[contains(@id, 'swal2-title') and (contains(text(), 'realizado com sucesso') or contains(text(), 'sucesso'))]")
                            self.log(f"Sucesso: {file_name} enviado.")
                            self.wait.until(EC.invisibility_of_element(popup_container))
                        except Exception as e_success: 
                            error_msg = "Mensagem de erro não encontrada"
                            try:
                                error_msg = popup_container.find_element(By.ID, "swal2-content").text
                            except:
                                 try: error_msg = popup_container.find_element(By.ID, "swal2-title").text
                                 except: pass
                            self.log(f"!!! ERRO DE UPLOAD (via popZ): {error_msg}")
                            try: popup_container.find_element(By.CLASS_NAME, "swal2-confirm").click()
                            except: pass 
                            self.wait.until(EC.invisibility_of_element(popup_container))
                            raise Exception(f"Erro de validação da página: {error_msg}")

                    except Exception as e_file:
                        self.log(f"!!! ERRO ao enviar o arquivo {file_name}: {e_file}")
                        self.log("...Tentando recarregar a página e continuar...")
                        self.driver.refresh()
                        self.log("Página recarregada. Aguardando para continuar...")
                        time.sleep(3) 
                        continue 

                if not self.stop_requested.is_set():
                    self.log("\n" + "="*50)
                    self.log("Upload de todos os arquivos concluído!")
                    self.log("="*50)
        
        except Exception as e:
            self.log(f"ERRO na Etapa 2: {e}")
        finally:
            self.log("Automação finalizada. Interface pronta para nova execução.")
            self.start_button.config(state="normal")
            self.captcha_button.config(state="disabled") 
            self.restart_button.config(state="normal") 
            self.stop_button.config(state="disabled") 
            self.stop_requested.clear() 


if __name__ == "__main__":
    root = ttk.Window(themename="superhero") 
    app = AutomationApp(root)
    root.mainloop()