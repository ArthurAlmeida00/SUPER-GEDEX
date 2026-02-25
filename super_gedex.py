import pandas as pd
import traceback
import time
import sys
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, StaleElementReferenceException

# ==============================================================================
# ### ÁREA DE CONFIGURAÇÃO DO USUÁRIO ###
# ==============================================================================

# 1. Login e Senha
MINHA_MATRICULA = "SEU_LOGIN"
MINHA_SENHA     = "SUA_SENHA"

# 2. Código da Subestação (O que será digitado na busca principal)
CODIGO_SUBESTACAO = "xxxxx"  # Mude aqui para 22245 ou qualquer outro

# 3. Filtro de Projeto (Coluna 5)
# Descomente (tire o #) da linha que você quer usar e comente as outras:
FILTRO_PROJETO = "SE_ELTC"
# FILTRO_PROJETO = "EQPRIMÁRIOS"
# FILTRO_PROJETO = "EQSECUNDÁRIOS"
# FILTRO_PROJETO = "TRANSF"
# FILTRO_PROJETO = ""  # Deixe vazio aspas duplas "" se não quiser filtrar nada nessa coluna

# 4. Outras Configurações
SALVAR_A_CADA = 20   # Salva backup a cada X linhas

# ==============================================================================

def iniciar_driver():
    print("Iniciando Driver...")
    # O ChromeDriverManager é ótimo para executáveis pois baixa o driver correto automaticamente
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service)
    driver.maximize_window()
    return driver

def fazer_login(driver):
    print("Acessando sistema...")
    driver.get("https://ged.cemig.com.br/Account/Login")
    
    try:
        driver.find_element(By.ID, "Login").send_keys(MINHA_MATRICULA)
        driver.find_element(By.ID, "Senha").send_keys(MINHA_SENHA)
    except: pass
    
    print("\n" + "="*50)
    print(">>> AÇÃO NECESSÁRIA: RESOLVA O CAPTCHA E ENTRE <<<")
    print("O robô está aguardando você entrar no sistema...")
    print("="*50 + "\n")

    # Espera a URL mudar (Sair do Login)
    try:
        WebDriverWait(driver, 600).until( # Aumentei para 10 min de tolerância
            lambda d: "Account/Login" not in d.current_url
        )
        print("\n>>> LOGIN DETECTADO! Iniciando automação...\n")
        time.sleep(2)
    except TimeoutException:
        print("Tempo esgotado! O login não foi detectado.")
        sys.exit() # Encerra o programa

def configurar_pesquisa_e_filtros(driver, cod_subestacao, filtro_proj):
    print("Navegando para Pesquisa Avançada...")
    driver.get("https://ged.cemig.com.br/Ficha/PesquisaAvancada")
    
    # Espera formulário
    WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "MaximoARecuperar")))

    try:
        print(f"Configurando busca para Subestação: {cod_subestacao}...")
        
        # Checkbox Paginação
        try:
            chk = driver.find_element(By.ID, "RemoverPaginacao")
            if not chk.is_selected(): chk.click()
        except: pass

        # Limite
        driver.find_element(By.ID, "MaximoARecuperar").clear()
        driver.find_element(By.ID, "MaximoARecuperar").send_keys("1000")

        # Código (Variável)
        cod = driver.find_element(By.XPATH, "//input[@placeholder='Código ou Nome da Aplicação']")
        cod.clear()
        cod.send_keys(cod_subestacao)
        time.sleep(2)
        cod.send_keys(Keys.TAB)

        # Clicar Pesquisar
        print("Clicando em Pesquisar...")
        driver.find_element(By.ID, "btnPesquisarFicha").click()
        
        # Espera TABELA
        print("Aguardando tabela aparecer...")
        WebDriverWait(driver, 30).until(EC.presence_of_element_located((By.XPATH, "//tr[.//span[@title='WORKFLOW']]")))
        time.sleep(3)

        # APLICAR FILTROS DE COLUNA
        print("Aplicando filtros de coluna...")
        
        # 1. Filtro Fixo: supergedex (Coluna 9)
        # Se quiser parametrizar esse também, me avise
        try:
            f1 = driver.find_element(By.CSS_SELECTOR, "input[data-index='9']")
            f1.clear(); f1.send_keys("supergedex"); time.sleep(0.5); f1.send_keys(Keys.ENTER)
            time.sleep(1.5)
        except: print("Aviso: Erro ao filtrar supergedex")

        # 2. Filtro Variável: PROJETO (Coluna 5)
        if filtro_proj and filtro_proj.strip() != "":
            print(f" -> Filtrando coluna 5 por: {filtro_proj}")
            try:
                f2 = driver.find_element(By.CSS_SELECTOR, "input[data-index='5']")
                f2.clear(); f2.send_keys(filtro_proj); time.sleep(0.5); f2.send_keys(Keys.ENTER)
                time.sleep(1.5)
            except: print(f"Aviso: Erro ao filtrar {filtro_proj}")
        else:
            print(" -> Nenhum filtro de projeto selecionado (Coluna 5 vazia).")
        
        print("Configuração concluída!")

    except Exception as e:
        print(f"Erro na configuração da pesquisa: {e}")

def extrair_dados_blindado(driver, deve_baixar_descricao):
    print("\nINICIANDO EXTRAÇÃO...")
    if not deve_baixar_descricao:
        print(" -> MODO RÁPIDO ATIVADO: Descrições serão ignoradas.")
    
    dados = []
    
    linhas_xpath = "//tr[.//span[@title='WORKFLOW']]"
    linhas = driver.find_elements(By.XPATH, linhas_xpath)
    total = len(linhas)
    print(f"Total de itens encontrados: {total}")

    for i in range(total):
        try:
            print(f"Processando {i+1}/{total}...", end="\r")
            
            # 1. Recaptura
            linha_atual = driver.find_elements(By.XPATH, linhas_xpath)[i]
            btn = linha_atual.find_element(By.CSS_SELECTOR, "span[title='WORKFLOW']")
            
            # 2. Clique Blindado
            modal_abriu = False
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
            time.sleep(1.0) 

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

            if not modal_abriu:
                print(f"\n[ERRO CRÍTICO] Modal não abriu na linha {i+1}")
                continue

            # 3. Extração
            time.sleep(1.5)
            modal = driver.find_element(By.ID, "modalWorkflowFichaCorpo")
            
            codigo = modal.get_attribute("data-name-ficha")
            versao = modal.get_attribute("data-rev-ficha")
            
            # --- LÓGICA DA PERGUNTA INICIAL ---
            descricao = "-"
            if deve_baixar_descricao:
                try: 
                    descricao = driver.find_element(By.ID, "modalWorkflowFichaTituloTitulo").text.strip()
                except: pass
            else:
                descricao = "Não Solicitado"

            # Status
            status = "INDEFINIDO"
            detalhe = "-"
            try:
                html_modal = modal.get_attribute("innerHTML").upper()
                if "CONCLUÍDO" in html_modal and "LABEL-XLG" in html_modal:
                    status = "APROVADO"
                    detalhe = "Finalizado"
                else:
                    tags = modal.find_elements(By.CSS_SELECTOR, ".timeline-item .widget-toolbar .label-white")
                    if tags:
                        txt = tags[0].text.upper()
                        detalhe = txt
                        if "EDIÇÃO" in txt: status = "REPROVADO"
                        elif "APROVAÇÃO" in txt: status = "EM ANÁLISE"
                        else: status = f"EM ANDAMENTO ({txt})"
                    else:
                        status = "EM ANDAMENTO"
            except: status = "ERRO"

            # Datas
            dt_att = "-"
            try:
                dts = modal.find_elements(By.CSS_SELECTOR, ".timeline-label b")
                if dts: dt_att = dts[0].text
            except: pass
            
            dt_ini = "-"
            try: dt_ini = modal.find_element(By.XPATH, ".//small[contains(., 'iniciado em')]//b").text
            except: pass

            dados.append({
                "Código": codigo,
                "Versão": versao,
                "Descrição": descricao,
                "Status": status,
                "Detalhe": detalhe,
                "Data Início": dt_ini,
                "Última Atualização": dt_att
            })

            # Fechar
            webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            try:
                WebDriverWait(driver, 3).until(EC.invisibility_of_element_located((By.ID, "modalWorkflowFichaCorpo")))
            except:
                time.sleep(1)

            # Backup
            if (i + 1) % SALVAR_A_CADA == 0:
                salvar_excel(dados, "Relatorio_Parcial.xlsx")

        except Exception as e:
            print(f"\n[ERRO REAL NA LINHA {i+1}]")
            traceback.print_exc()
            try: webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            except: pass

    return dados

def salvar_excel(dados, nome_arquivo):
    if not dados: return
    df = pd.DataFrame(dados)
    cols = ["Código", "Versão", "Status", "Detalhe", "Data Início", "Última Atualização", "Descrição"]
    cols_fin = [c for c in cols if c in df.columns]
    df = df[cols_fin]
    df.to_excel(nome_arquivo, index=False)

# ==============================================================================
# EXECUÇÃO PRINCIPAL
# ==============================================================================
if __name__ == "__main__":
    try:
        print("="*50)
        print("   ROBÔ GEDEX AUTOMATION - Versão 2.0")
        print("="*50)
        
        # 1. Pergunta interativa inicial
        resp_usuario = input("\n>> Deseja baixar a DESCRIÇÃO dos itens? (S/N): ").strip().upper()
        BAIXAR_DESCRICAO = (resp_usuario == 'S')
        
        print(f"\nConfiguração Atual:")
        print(f" - Subestação: {CODIGO_SUBESTACAO}")
        print(f" - Filtro Projeto: {FILTRO_PROJETO}")
        print(f" - Baixar Descrição: {'SIM' if BAIXAR_DESCRICAO else 'NÃO'}")
        print("-" * 30)

        driver = iniciar_driver()
        
        fazer_login(driver)
        
        # Passamos os parâmetros configurados lá em cima
        configurar_pesquisa_e_filtros(driver, CODIGO_SUBESTACAO, FILTRO_PROJETO)
        
        # Passamos a escolha do usuário sobre a descrição
        dados_finais = extrair_dados_blindado(driver, BAIXAR_DESCRICAO)
        
        print("\n" + "="*40)
        nome_final = f"Relatorio_GEDEX_{CODIGO_SUBESTACAO}_{FILTRO_PROJETO}.xlsx"
        salvar_excel(dados_finais, nome_final)
        print(f"SUCESSO TOTAL! Arquivo salvo como: {nome_final}")
        
    except Exception as e:
        print("\nERRO CRÍTICO NO PROGRAMA:")
        traceback.print_exc()
    finally:
        print("\nProcesso finalizado.")
        # O input abaixo impede que a janela feche sozinha no executável
        input("Pressione ENTER para fechar a janela...")
        try: driver.quit()
        except: pass