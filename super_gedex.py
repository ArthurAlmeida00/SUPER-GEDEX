import tkinter as tk
from tkinter import ttk, messagebox, PhotoImage
import pandas as pd
import traceback
import time
import sys
import os
import threading  # <--- IMPORTANTE: Permite rodar o robô sem travar a tela
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ==============================================================================
# LÓGICA DO ROBÔ (Adaptada para atualizar a Barra)
# ==============================================================================

def executar_robo(dados_interface, progress_bar, lbl_status, btn_iniciar, root):
    LOGIN = dados_interface['login']
    SENHA = dados_interface['senha']
    SUBESTACAO = dados_interface['subestacao']
    FILTRO_PROJETO = dados_interface['filtro']
    BAIXAR_DESCRICAO = dados_interface['baixar_descricao']
    BAIXAR_DOCS = dados_interface['baixar_docs']
    SALVAR_A_CADA = 20

    # Função auxiliar para atualizar texto na tela (thread-safe)
    def atualizar_status(texto, valor_barra=None, max_barra=None):
        lbl_status.config(text=texto)
        if max_barra is not None:
            progress_bar['maximum'] = max_barra
        if valor_barra is not None:
            progress_bar['value'] = valor_barra
        root.update_idletasks() # Força o desenho da tela

    try:
        atualizar_status("Iniciando Navegador...", 0, 100)
        
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service)
        driver.maximize_window()

        # --- LOGIN ---
        atualizar_status("Acessando Login...", 5)
        driver.get("https://ged.cemig.com.br/Account/Login")
        try:
            driver.find_element(By.ID, "Login").send_keys(LOGIN)
            driver.find_element(By.ID, "Senha").send_keys(SENHA)
        except: pass
        
        atualizar_status("AGUARDANDO LOGIN MANUAL (Resolva o Captcha)...", 10)
        
        # Espera Login
        try:
            WebDriverWait(driver, 600).until(lambda d: "Account/Login" not in d.current_url)
            atualizar_status("Login Detectado! Configurando busca...", 15)
            time.sleep(2)
        except TimeoutException:
            messagebox.showerror("Erro", "Tempo de login esgotado.")
            driver.quit()
            btn_iniciar.config(state="normal") # Reabilita botão
            return

        # --- BUSCA ---
        driver.get("https://ged.cemig.com.br/Ficha/PesquisaAvancada")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "MaximoARecuperar")))

        try:
            chk = driver.find_element(By.ID, "RemoverPaginacao")
            if not chk.is_selected(): chk.click()
        except: pass

        driver.find_element(By.ID, "MaximoARecuperar").clear()
        driver.find_element(By.ID, "MaximoARecuperar").send_keys("1000")

        cod = driver.find_element(By.XPATH, "//input[@placeholder='Código ou Nome da Aplicação']")
        cod.clear()
        cod.send_keys(SUBESTACAO)
        time.sleep(2)
        cod.send_keys(Keys.TAB)

        driver.find_element(By.ID, "btnPesquisarFicha").click()
        
        atualizar_status("Aguardando carregamento da tabela...", 25)
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//tr[.//span[@title='WORKFLOW']]")))
        time.sleep(3)

        # --- FILTROS ---
        atualizar_status("Aplicando filtros de coluna...", 30)
        try:
            f1 = driver.find_element(By.CSS_SELECTOR, "input[data-index='9']")
            f1.clear(); f1.send_keys("supergedex"); time.sleep(0.5); f1.send_keys(Keys.ENTER)
            time.sleep(1.5)
        except: pass

        if FILTRO_PROJETO != "TODOS":
            try:
                f2 = driver.find_element(By.CSS_SELECTOR, "input[data-index='5']")
                f2.clear(); f2.send_keys(FILTRO_PROJETO); time.sleep(0.5); f2.send_keys(Keys.ENTER)
                time.sleep(1.5)
            except: pass

        # --- PREPARAÇÃO DA EXTRAÇÃO ---
        dados = []
        linhas_xpath = "//tr[.//span[@title='WORKFLOW']]"
        linhas = driver.find_elements(By.XPATH, linhas_xpath)
        total = len(linhas)
        
        # Configura a barra para o tamanho exato da lista
        atualizar_status(f"Iniciando: 0/{total}", 0, total)
        print(f"Total encontrado: {total}")

        for i in range(total):
            try:
                # Atualiza Interface Visual
                atualizar_status(f"Processando: {i+1}/{total}", i+1)
                
                # Lógica de Extração
                linha_atual = driver.find_elements(By.XPATH, linhas_xpath)[i]
                btn = linha_atual.find_element(By.CSS_SELECTOR, "span[title='WORKFLOW']")
                
                driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
                time.sleep(1.0)
                
                modal_abriu = False
                try:
                    driver.execute_script("arguments[0].click();", btn)
                    try:
                        WebDriverWait(driver, 1.0).until(EC.visibility_of_element_located((By.ID, "modalWorkflowFichaCorpo")))
                        modal_abriu = True
                    except:
                        driver.execute_script("arguments[0].click();", btn)
                        WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "modalWorkflowFichaCorpo")))
                        modal_abriu = True
                except: pass

                if not modal_abriu: continue

                time.sleep(1.5)
                modal = driver.find_element(By.ID, "modalWorkflowFichaCorpo")
                codigo = modal.get_attribute("data-name-ficha")
                versao = modal.get_attribute("data-rev-ficha")

                descricao = "-"
                if BAIXAR_DESCRICAO:
                    try: descricao = driver.find_element(By.ID, "modalWorkflowFichaTituloTitulo").text.strip()
                    except: pass
                else: descricao = "Não Solicitado"

                status = "INDEFINIDO"; detalhe = "-"
                try:
                    html_modal = modal.get_attribute("innerHTML").upper()
                    if "CONCLUÍDO" in html_modal and "LABEL-XLG" in html_modal:
                        status = "APROVADO"; detalhe = "Finalizado"
                    else:
                        tags = modal.find_elements(By.CSS_SELECTOR, ".timeline-item .widget-toolbar .label-white")
                        if tags:
                            txt = tags[0].text.upper(); detalhe = txt
                            if "EDIÇÃO" in txt: status = "REPROVADO"
                            elif "APROVAÇÃO" in txt: status = "EM ANÁLISE"
                            else: status = f"EM ANDAMENTO ({txt})"
                        else: status = "EM ANDAMENTO"
                except: pass

                dt_ini = "-"; dt_att = "-"
                try: dt_ini = modal.find_element(By.XPATH, ".//small[contains(., 'iniciado em')]//b").text
                except: pass
                try:
                    dts = modal.find_elements(By.CSS_SELECTOR, ".timeline-label b")
                    if dts: dt_att = dts[0].text
                except: pass

                msg_docs = "Não Solicitado"
                if BAIXAR_DOCS: msg_docs = "Em Breve"

                dados.append({
                    "Código": codigo, "Versão": versao, "Descrição": descricao,
                    "Status": status, "Detalhe": detalhe, 
                    "Data Início": dt_ini, "Última Atualização": dt_att,
                    "Documentos": msg_docs
                })

                webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                try: WebDriverWait(driver, 3).until(EC.invisibility_of_element_located((By.ID, "modalWorkflowFichaCorpo")))
                except: time.sleep(1)

                if (i + 1) % SALVAR_A_CADA == 0:
                    pd.DataFrame(dados).to_excel("Relatorio_Parcial.xlsx", index=False)

            except Exception as e:
                print(f"Erro linha {i+1}: {e}")
                try: webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                except: pass

        if dados:
            nome_arq = f"Relatorio_{SUBESTACAO}_{FILTRO_PROJETO}.xlsx"
            pd.DataFrame(dados).to_excel(nome_arq, index=False)
            atualizar_status("Concluído!")
            messagebox.showinfo("Sucesso", f"Processo finalizado!\nArquivo: {nome_arq}")
        else:
            atualizar_status("Nenhum dado encontrado.")
            messagebox.showwarning("Aviso", "Nenhum dado foi coletado.")
        
        driver.quit()

    except Exception as e:
        traceback.print_exc()
        messagebox.showerror("Erro Fatal", f"Erro: {e}")
    finally:
        # Reabilita o botão Iniciar no final de tudo
        btn_iniciar.config(state="normal")

# ==============================================================================
# INTERFACE GRÁFICA
# ==============================================================================

def iniciar_thread():
    # Coleta dados
    dados = {
        'login': entry_login.get(),
        'senha': entry_senha.get(),
        'subestacao': entry_sub.get(),
        'filtro': combo_filtro.get(),
        'baixar_descricao': var_desc.get(),
        'baixar_docs': var_docs.get()
    }

    if not dados['login'] or not dados['senha'] or not dados['subestacao']:
        messagebox.showwarning("Atenção", "Preencha Login, Senha e Subestação!")
        return

    # Trava o botão para não clicar duas vezes
    btn_iniciar.config(state="disabled")
    
    # Mostra a barra de progresso
    frame_status.pack(pady=10, fill="x", padx=10)
    progress_bar['value'] = 0
    lbl_status.config(text="Preparando...")

    # Cria uma Thread separada para o robô não travar a janela
    t = threading.Thread(target=executar_robo, args=(dados, progress_bar, lbl_status, btn_iniciar, root))
    t.start()

# Configuração da Janela
root = tk.Tk()
root.title("Versão 1.0 - Loading")
root.geometry("420x700") # Mais altura para a barra
style = ttk.Style()
style.theme_use('clam')

# Logo
def resource_path(relative_path):
    """ Retorna o caminho absoluto para o recurso, funciona no dev e no PyInstaller """
    try:
        
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
try:
    caminho_logo = resource_path("logo.png") # <--- Usa a função mágica
    if os.path.exists(caminho_logo):
        img_logo = PhotoImage(file=caminho_logo)
        lbl_img = tk.Label(root, image=img_logo)
        lbl_img.pack(pady=10)
    else:
        tk.Label(root, text="[Logo não encontrado]", fg="gray").pack(pady=5)
except Exception as e:
    print(f"Erro ao carregar imagem: {e}")

tk.Label(root, text="Automação GEDEX", font=("Arial", 16, "bold"), pady=5).pack()

# Login
frame_login = ttk.LabelFrame(root, text="Credenciais")
frame_login.pack(padx=10, pady=5, fill="x")
ttk.Label(frame_login, text="Login:").pack(anchor="w", padx=5)
entry_login = ttk.Entry(frame_login); entry_login.pack(fill="x", padx=5, pady=2)
ttk.Label(frame_login, text="Senha:").pack(anchor="w", padx=5)
entry_senha = ttk.Entry(frame_login, show="*"); entry_senha.pack(fill="x", padx=5, pady=2)

# Busca
frame_busca = ttk.LabelFrame(root, text="Configuração")
frame_busca.pack(padx=10, pady=5, fill="x")
ttk.Label(frame_busca, text="Código Subestação:").pack(anchor="w", padx=5)
entry_sub = ttk.Entry(frame_busca); entry_sub.pack(fill="x", padx=5, pady=2)
entry_sub.insert(0, "xxxxx")
ttk.Label(frame_busca, text="Filtro Projeto:").pack(anchor="w", padx=5)
combo_filtro = ttk.Combobox(frame_busca, values=["TODOS", "SE_ELTC", "EQPRIMÁRIOS", "EQSECUNDÁRIOS", "TRANSF"], state="readonly")
combo_filtro.current(0)
combo_filtro.pack(fill="x", padx=5, pady=2)

# Opções
frame_ops = ttk.LabelFrame(root, text="Opções")
frame_ops.pack(padx=10, pady=5, fill="x")
var_desc = tk.BooleanVar(value=True)
ttk.Checkbutton(frame_ops, text="Baixar Descrição", variable=var_desc).pack(anchor="w", padx=5)
var_docs = tk.BooleanVar(value=False)
ttk.Checkbutton(frame_ops, text="Baixar Documentos (Em breve)", variable=var_docs).pack(anchor="w", padx=5)

# --- ÁREA DE STATUS (BARRA DE LOADING) ---
frame_status = ttk.LabelFrame(root, text="Status do Processo")


lbl_status = tk.Label(frame_status, text="Aguardando início...", font=("Arial", 10))
lbl_status.pack(pady=5)

# Barra de Progresso Verde
style.configure("green.Horizontal.TProgressbar", foreground='green', background='green')
progress_bar = ttk.Progressbar(frame_status, style="green.Horizontal.TProgressbar", orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=5, padx=10, fill="x")

# Botão
btn_iniciar = tk.Button(root, text="INICIAR", font=("Arial", 12, "bold"), bg="#4CAF50", fg="white", height=2, command=iniciar_thread)
btn_iniciar.pack(padx=10, pady=20, fill="x")

# Rodapé
tk.Label(root, text="Última atualização por: Arthur Almeida", font=("Arial", 8, "italic"), fg="gray").pack(side=tk.BOTTOM, pady=10)

root.mainloop()