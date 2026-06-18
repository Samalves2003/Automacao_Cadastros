import os
import sys
import time
import subprocess
import threading
import pandas as pd
import customtkinter as ctk
from tkinter import filedialog, messagebox

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# Configurações de aparência do CustomTkinter
ctk.set_appearance_mode("System")  # Segue o modo do Windows (Dark/Light)
ctk.set_default_color_theme("blue")

class BotExcelSalesforceApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Configurações da Janela
        self.title("Automação de Cadastros - Salesforce (FOCO)")
        self.geometry("850x650")
        self.minsize(800, 600)

        # Variáveis de Controle de Estado
        self.navegador_aberto = False
        self.excel_selecionado = False
        self.caminho_excel = ""
        self.driver = None

        # --- ESTRUTURA LAYOUT (Grid) ---
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # 1. Cabeçalho / Título
        self.lbl_titulo = ctk.CTkLabel(self, text="Painel de Controle do Bot", font=ctk.CTkFont(size=22, weight="bold"))
        self.lbl_titulo.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")

        # 2. Área de Instruções e Logs
        self.frame_conteudo = ctk.CTkFrame(self)
        self.frame_conteudo.grid(row=1, column=0, padx=20, pady=10, sticky="nsew")
        self.frame_conteudo.grid_columnconfigure(0, weight=1)
        self.frame_conteudo.grid_columnconfigure(1, weight=1)
        self.frame_conteudo.grid_rowconfigure(0, weight=1)

        self.txt_instrucoes = ctk.CTkTextbox(self.frame_conteudo, wrap="word", font=ctk.CTkFont(size=13))
        self.txt_instrucoes.grid(row=0, column=0, padx=(10, 5), pady=10, sticky="nsew")
        self.carregar_instrucoes()

        self.txt_logs = ctk.CTkTextbox(self.frame_conteudo, wrap="word", font=ctk.CTkFont(family="Courier", size=12))
        self.txt_logs.grid(row=0, column=1, padx=(5, 10), pady=10, sticky="nsew")
        self.log("[SISTEMA] Aguardando inicialização do ambiente...")

        # 3. Frame de Controles
        self.frame_botoes = ctk.CTkFrame(self, height=120)
        self.frame_botoes.grid(row=2, column=0, padx=20, pady=(10, 20), sticky="ew")
        self.frame_botoes.grid_columnconfigure((0, 1, 2), weight=1)

        self.btn_navegador = ctk.CTkButton(self.frame_botoes, text="1. Abrir Navegador", command=self.acao_abrir_navegador, height=45)
        self.btn_navegador.grid(row=0, column=0, padx=10, pady=15, sticky="ew")

        self.btn_selecionar_excel = ctk.CTkButton(self.frame_botoes, text="2. Selecionar Planilha", command=self.acao_selecionar_excel, height=45)
        self.btn_selecionar_excel.grid(row=0, column=1, padx=10, pady=15, sticky="ew")

        self.btn_rodar = ctk.CTkButton(self.frame_botoes, text="3. Rodar Bot", command=self.acao_rodar_bot, state="disabled", fg_color="gray", height=45)
        self.btn_rodar.grid(row=0, column=2, padx=10, pady=15, sticky="ew")

        self.lbl_status_arquivo = ctk.CTkLabel(self.frame_botoes, text="Nenhum arquivo selecionado", font=ctk.CTkFont(size=11), text_color="gray")
        self.lbl_status_arquivo.grid(row=1, column=0, columnspan=3, pady=(0, 10))

    def carregar_instrucoes(self):
        instrucoes = (
            "=== INSTRUÇÕES DO SISTEMA ===\n\n"
            "Do que se trata o Bot?\n"
            "Este automatizador atualiza dados cadastrais (E-mail, Telefone, Canal Preferencial, Endereço e Consentimentos) no Salesforce.\n\n"
            "Colunas obrigatórias no Excel:\n"
            "• CPF : Apenas números ou com pontos e traço.\n"
            "• Telefone : Número do telefone.\n"
            "• Tipo de telefone : (ex: 'Telefone Celular', 'Telefone Comercial').\n"
            "• Email : Endereço de e-mail.\n"
            "• WPP / ARE / ARL / ARM : 'SIM' para autorizar.\n"
            "• Canal preferencial : Opções (ex: 'Email', 'WhatsApp', 'SMS').\n\n"
            "Passo a Passo:\n"
            "1. Clique em '1. Abrir Navegador'.\n"
            "2. Faça seu Login no FOCO na janela que abrir.\n"
            "3. Clique em '2. Selecionar Planilha'.\n"
            "4. Clique em '3. Rodar Bot'."
        )
        self.txt_instrucoes.insert("0.0", instrucoes)
        self.txt_instrucoes.configure(state="disabled")

    def log(self, messaging):
        self.txt_logs.configure(state="normal")
        self.txt_logs.insert("end", f"{messaging}\n")
        self.txt_logs.see("end")
        self.txt_logs.configure(state="disabled")

    def atualizar_estado_botao_rodar(self):
        if self.navegador_aberto and self.excel_selecionado:
            self.btn_rodar.configure(state="normal", fg_color=("#2b719e", "#1f538d"))
            self.log("[SISTEMA] Bot pronto para ser iniciado.")
        else:
            self.btn_rodar.configure(state="disabled", fg_color="gray")

    def acao_abrir_navegador(self):
        threading.Thread(target=self._thread_iniciar_chrome, daemon=True).start()

    def _thread_iniciar_chrome(self):
        self.log("[NAVEGADOR] Abrindo o Google Chrome em modo de Depuração...")
        chrome_cmd = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
        user_dir = r'C:\ChromeDevSession'
        cmd = f'"{chrome_cmd}" --remote-debugging-port=9222 --user-data-dir="{user_dir}"'
        
        try:
            subprocess.Popen(cmd, shell=True)
            self.log("[NAVEGADOR] Chrome aberto com sucesso!")
            self.log("[AVISO] Faça o Login no FOCO antes de Prosseguir.")
            
            self.navegador_aberto = True
            self.atualizar_estado_botao_rodar()
        except Exception as e:
            self.log(f"[ERRO] Falha ao abrir o Chrome: {e}")
            messagebox.showerror("Erro", f"Falha ao abrir o Chrome:\n{e}")

    def acao_selecionar_excel(self):
        caminho = filedialog.askopenfilename(title="Selecione a base", filetypes=[("Excel", "*.xlsx *.xls")])
        if caminho:
            self.caminho_excel = caminho
            nome_arquivo = os.path.basename(caminho)
            self.lbl_status_arquivo.configure(text=f"Arquivo: {nome_arquivo}", text_color="green")
            self.log(f"[ARQUIVO] Planilha selecionada: {nome_arquivo}")
            self.excel_selecionado = True
            self.atualizar_estado_botao_rodar()

    def acao_rodar_bot(self):
        self.btn_rodar.configure(state="disabled", text="Processando...")
        self.btn_selecionar_excel.configure(state="disabled")
        self.btn_navegador.configure(state="disabled")
        threading.Thread(target=self._thread_loop_principal, daemon=True).start()

    def _thread_loop_principal(self):
        self.log("\n[EXECUÇÃO] Conectando ao navegador...")
        
        try:
            opts = Options()
            opts.add_experimental_option("debuggerAddress", "127.0.0.1:9222")
            self.driver = webdriver.Chrome(options=opts)
            
            if len(self.driver.window_handles) > 0:
                self.driver.switch_to.window(self.driver.window_handles[0])
        except Exception as e:
            self.log(f"[ERRO] Não foi possível conectar ao Chrome: {e}")
            self.finalizar_processamento()
            return

        self.log("[EXECUÇÃO] Lendo a planilha...")
        try:
            df = pd.read_excel(self.caminho_excel, dtype=str)
        except Exception as e:
            self.log(f"[ERRO] Falha ao ler Excel: {e}")
            self.finalizar_processamento()
            return

        resultados = []
        total_linhas = len(df)

        for index, row in df.iterrows():
            self.log(f"--- Registro {index + 1} de {total_linhas} ---")
            status = self.processar_linha_salesforce(self.driver, row)
            
            resultados.append({
                "CPF": row.get('CPF', 'N/A'), 
                "Telefone": row.get('Telefone', 'N/A'),
                "Tipo de telefone": row.get('Tipo de telefone', 'N/A'),
                "Email": row.get('Email', 'N/A'),
                "WPP": row.get('WPP', 'N/A'),
                "ARE": row.get('ARE', 'N/A'),
                "ARL": row.get('ARL', 'N/A'),
                "ARM": row.get('ARM', 'N/A'),
                "Canal preferencial": row.get('Canal preferencial', 'N/A'),
                "Status": status
            })

        try:
            pd.DataFrame(resultados).to_excel("Log_Final.xlsx", index=False)
            self.log("\n[FIM] Relatório 'Log_Final.xlsx' gerado.")
            messagebox.showinfo("Sucesso", "Processamento finalizado!")
        except Exception as e:
            self.log(f"[ERRO] Falha ao salvar Log: {e}")
            messagebox.showerror("Erro de Permissão", "Não foi possível salvar o Log_Final.xlsx.")
            
        self.finalizar_processamento()

    def finalizar_processamento(self):
        self.btn_selecionar_excel.configure(state="normal")
        self.btn_navegador.configure(state="normal")
        self.btn_rodar.configure(state="normal", text="3. Rodar Bot")

    def processar_linha_salesforce(self, driver, row_data):
        wait = WebDriverWait(driver, 20)
        
        # Extração Segura do CPF (Preservando a pontuação se existir)
        cpf_bruto = str(row_data.get('CPF', '')).strip() if pd.notna(row_data.get('CPF')) else ""
        if cpf_bruto.endswith(".0") and cpf_bruto.count(".") == 1:
            cpf = cpf_bruto[:-2]
        else:
            cpf = cpf_bruto

        # Extração Segura do Telefone
        tel_bruto = str(row_data.get('Telefone', '')).strip() if pd.notna(row_data.get('Telefone')) else ""
        if tel_bruto.endswith(".0") and tel_bruto.count(".") == 1:
            telefone = tel_bruto[:-2]
        else:
            telefone = tel_bruto

        email = str(row_data.get('Email', '')).strip() if pd.notna(row_data.get('Email')) else ""
        wpp = str(row_data.get('WPP', '')).strip() if pd.notna(row_data.get('WPP')) else ""
        are = str(row_data.get('ARE', '')).strip().upper() if pd.notna(row_data.get('ARE')) else ""
        arl = str(row_data.get('ARL', '')).strip().upper() if pd.notna(row_data.get('ARL')) else ""
        arm = str(row_data.get('ARM', '')).strip().upper() if pd.notna(row_data.get('ARM')) else ""
        tipo_telefone = str(row_data.get('Tipo de telefone', '')).strip() if pd.notna(row_data.get('Tipo de telefone')) else ""
        canal_preferencial = str(row_data.get('Canal preferencial', '')).strip() if pd.notna(row_data.get('Canal preferencial')) else ""

        self.log(f" -> CPF: {cpf}")

        try:
            def js_click(elm): driver.execute_script("arguments[0].click();", elm)
            def scroll_ate(elm): driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elm); time.sleep(0.5)

            driver.get("https://sebraecrm.lightning.force.com/lightning/o/Account/new?inContextOfRef=1.eyJ0eXBlIjoic3RhbmRhcmRfX29iamVjdFBhZ2UiLCJhdHRyaWJ1dGVzIjp7Im9iamVjdEFwaU5hbWUiOiJBY2NvdW50IiwiYWN0aW9uTmFtZSI6Imxpc3QifSwic3RhdGUiOnsiZmlsdGVyTmFtZSI6Il9fUmVjZW50In19&count=3")
            wait.until(EC.visibility_of_element_located((By.XPATH, "//label[text()='CPF']")))
            time.sleep(1)

            xpath_input_cpf = "//label[text()='CPF']/following::div/input[@class='slds-input']"
            input_cpf = wait.until(EC.presence_of_element_located((By.XPATH, xpath_input_cpf)))
            scroll_ate(input_cpf)
            js_click(input_cpf)
            input_cpf.send_keys(Keys.CONTROL + "a")
            input_cpf.send_keys(Keys.BACKSPACE)
            input_cpf.send_keys(cpf)
            time.sleep(0.5)

            xpath_btn_buscar_modal = "//button[contains(@class, 'slds-button_brand')][text()='Buscar']"
            js_click(wait.until(EC.presence_of_element_located((By.XPATH, xpath_btn_buscar_modal))))
            time.sleep(2) 
            
            js_click(wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='radio' and contains(@name, 'options')]"))))
            time.sleep(1)
            
            js_click(wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'slds-button_brand')][text()='Continuar']"))))
            time.sleep(2)

            # --- E-MAIL ---
            if email:
                xpath_input_email = "//label[text()='E-mail Principal']/following::div/input[@name='Email__c' or @class='slds-input']"
                input_email = wait.until(EC.presence_of_element_located((By.XPATH, xpath_input_email)))
                scroll_ate(input_email)
                js_click(input_email) 
                input_email.send_keys(Keys.CONTROL + "a")
                input_email.send_keys(Keys.BACKSPACE)
                input_email.send_keys(email)
                time.sleep(1)

                btn_combobox = wait.until(EC.presence_of_element_located((By.XPATH, "//label[text()='Autoriza Receber E-mail']/following::button[@role='combobox']")))
                js_click(btn_combobox)
                time.sleep(0.5)  

                if are == "SIM":
                    opcao_are = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@aria-label='Autoriza Receber E-mail']//lightning-base-combobox-item[@data-value='Sim']")))
                else:
                    opcao_are = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@aria-label='Autoriza Receber E-mail']//lightning-base-combobox-item[@data-value='Não']")))
                js_click(opcao_are)
                time.sleep(1)

            # --- TELEFONE ---
            if telefone:
                xpath_input_tel = "//label[text()='Telefone Principal']/following::div/input[@name='Phone' or @class='slds-input']"
                input_tel = wait.until(EC.presence_of_element_located((By.XPATH, xpath_input_tel)))
                scroll_ate(input_tel)  
                js_click(input_tel) 
                input_tel.send_keys(Keys.CONTROL + "a")
                input_tel.send_keys(Keys.BACKSPACE)
                input_tel.send_keys(telefone)
                time.sleep(1)

                btn_wpp = wait.until(EC.presence_of_element_located((By.XPATH, "//label[text()='Telefone Whatsapp?']/following::button[@role='combobox']")))
                scroll_ate(btn_wpp)
                js_click(btn_wpp)
                time.sleep(0.5)

                xpath_opcao_wpp = "//div[@aria-label='Telefone Whatsapp?']//lightning-base-combobox-item[@data-value='Telefone Principal']"
                js_click(wait.until(EC.presence_of_element_located((By.XPATH, xpath_opcao_wpp))))
                time.sleep(1)

                btn_combobox = wait.until(EC.presence_of_element_located((By.XPATH, "//label[text()='Autoriza Receber Ligação']/following::button[@role='combobox']")))
                js_click(btn_combobox)
                time.sleep(0.5)  

                if arl == "SIM":
                    opcao_arl = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@aria-label='Autoriza Receber Ligação']//lightning-base-combobox-item[@data-value='Sim']")))
                else:
                    opcao_arl = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@aria-label='Autoriza Receber Ligação']//lightning-base-combobox-item[@data-value='Não']")))
                js_click(opcao_arl)
                time.sleep(1)
                    
                btn_combobox = wait.until(EC.presence_of_element_located((By.XPATH, "//label[text()='Autoriza Receber Mensagem']/following::button[@role='combobox']")))
                js_click(btn_combobox)
                time.sleep(0.5)  

                if arm == "SIM":
                    opcao_arm = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@aria-label='Autoriza Receber Mensagem']//lightning-base-combobox-item[@data-value='Sim']")))
                else:
                    opcao_arm = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@aria-label='Autoriza Receber Mensagem']//lightning-base-combobox-item[@data-value='Não']")))
                js_click(opcao_arm)
                time.sleep(1)

                btn_combobox = wait.until(EC.presence_of_element_located((By.XPATH, "//label[text()='Autoriza Receber Whatsapp?']/following::button[@role='combobox']")))
                scroll_ate(btn_combobox)
                js_click(btn_combobox)
                time.sleep(0.5)  

                if wpp.upper() == "SIM":
                    opcao_wpp_auth = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@aria-label='Autoriza Receber Whatsapp?']//lightning-base-combobox-item[@data-value='Sim']")))
                else:
                    opcao_wpp_auth = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@aria-label='Autoriza Receber Whatsapp?']//lightning-base-combobox-item[@data-value='Não']")))
                js_click(opcao_wpp_auth)
                time.sleep(1)

                if tipo_telefone:
                    btn_tipo_tel = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@aria-label='Tipo Telefone Principal']")))
                    scroll_ate(btn_tipo_tel)
                    js_click(btn_tipo_tel)
                    time.sleep(0.5)

                    xpath_opcao_tipo = f"//div[@aria-label='Tipo Telefone Principal']//lightning-base-combobox-item[@data-value='{tipo_telefone}']"
                    js_click(wait.until(EC.presence_of_element_located((By.XPATH, xpath_opcao_tipo))))
                    time.sleep(1)

                btn_status_tel = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@aria-label='Status Telefone Principal']")))
                scroll_ate(btn_status_tel)
                js_click(btn_status_tel)
                time.sleep(0.5)

                xpath_opcao_status = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@aria-label='Status Telefone Principal']//lightning-base-combobox-item[@data-value='Ativo']")))
                js_click(xpath_opcao_status)
                time.sleep(1)

            # --- CANAL PREFERENCIAL ---
            if canal_preferencial:
                btn_canal_pref = wait.until(EC.presence_of_element_located((By.XPATH, "//button[@aria-label='Canal Preferencial']")))
                scroll_ate(btn_canal_pref)
                js_click(btn_canal_pref)
                time.sleep(0.5)

                xpath_opcao_canal = f"//div[@aria-label='Canal Preferencial']//lightning-base-combobox-item[@data-value='{canal_preferencial}']"
                js_click(wait.until(EC.presence_of_element_located((By.XPATH, xpath_opcao_canal))))
                time.sleep(1)

            # --- SALVAMENTO FINAL COM TENTATIVA DUPLA ---
            xpath_btn_salvar = "//button[@name='SaveEdit'][text()='Salvar']"
            btn_salvar = wait.until(EC.presence_of_element_located((By.XPATH, xpath_btn_salvar)))
            scroll_ate(btn_salvar)
            
            # 1ª Tentativa de clique
            js_click(btn_salvar)
            
            try:
                WebDriverWait(driver, 4).until(EC.invisibility_of_element_located((By.XPATH, xpath_btn_salvar)))
            except:
                self.log(" [AVISO] O modal não fechou no primeiro clique. Tentando salvar novamente (2ª vez)...")
                js_click(btn_salvar)
                
                try:
                    WebDriverWait(driver, 4).until(EC.invisibility_of_element_located((By.XPATH, xpath_btn_salvar)))
                except:
                    self.log(" [ERRO] O registro permaneceu bloqueado na tela de salvamento. Pulando registro...")
                    return "Erro - Bloqueado na tela de salvamento (Análise humana necessária)"

            time.sleep(1.5)
            self.log(" [OK] Processado com sucesso.")
            return "Processado com sucesso"

        except Exception as e:
            self.log(f" [ERRO] {e}")
            return "Erro - Necessita de análise humana"

if __name__ == "__main__":
    app = BotExcelSalesforceApp()
    app.mainloop()