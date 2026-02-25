from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By 
import time
import traceback
from selenium.webdriver.common.keys import Keys
import pandas as pd 
import re
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

# ---  DADOS ---

MINHA_MATRICULA = "LOGIN"
MINHA_SENHA = "Cemig"

# 1. Configuração Inicial
print("Iniciando o Robô...")
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service)

# 2. Ação: Entra no site
link_do_favorito = "https://ged.cemig.com.br/Account/Login" 

print(f"Indo para: {link_do_favorito}")
driver.get(link_do_favorito)

# 3. Ação: Preencher Login (A parte nova)
try:
    print("Tentando preencher credenciais...")
    
    # Procura o campo com ID="Login" e escreve a matrícula
    driver.find_element(By.ID, "Login").send_keys(MINHA_MATRICULA)
    
    # Procura o campo com ID="Senha" e escreve a senha
    driver.find_element(By.ID, "Senha").send_keys(MINHA_SENHA)
    
    print(">>> Sucesso! Matrícula e Senha preenchidas.")
    
except Exception as e:
    print(f"Erro ao tentar preencher: {e}")

# 4. Pausa para o Humano (Captcha)
print("\n" + "="*40)
print("LOGUE")
print("="*40 + "\n")

input("SE LOGADO PRESSIONE ENTER")

print("Logado! Navegando para Pesquisa Avançada...")
url_pesquisa = "https://ged.cemig.com.br/Ficha/PesquisaAvancada"
driver.get(url_pesquisa)

# --- CORREÇÃO 1: Espera a página carregar antes de tentar preencher ---
time.sleep(3) 

print("\n" + "="*40)
print("PREENCHENDO FILTROS SOZINHO...")

try:
    # 1. Checkbox "Remover Paginação"
    try:
        checkbox = driver.find_element(By.ID, "RemoverPaginacao")
        if not checkbox.is_selected():
            checkbox.click()
            print(" -> Checkbox 'Remover Paginação' marcado.")
    except:
        pass 

    # 2. Campo Limite
    print(" -> Definindo limite para 1000...")
    campo_limite = driver.find_element(By.ID, "MaximoARecuperar")
    campo_limite.clear()
    campo_limite.send_keys("1000")

    # 3. Campo Código 
    print(" -> Digitando código XXXXX e validando...")
    campo_codigo = driver.find_element(By.XPATH, "//input[@placeholder='Código ou Nome da Aplicação']")
    campo_codigo.clear()
    campo_codigo.send_keys("XXXXX")
    
    # PAUSA CRÍTICA: Espera 2 segundos para o menu azul aparecer "embaixo" do texto
    time.sleep(2) 
    
    # Pressiona TAB para selecionar a opção que apareceu
    print(" -> Pressionando TAB para confirmar...")
    campo_codigo.send_keys(Keys.TAB)
    
except Exception as e:
    print(f"ERRO AO PREENCHER FILTROS: {e}")

# 6. Clicar no botão PESQUISAR e Iniciar Fluxo Contínuo
try:
    print("O Robô está clicando no botão Pesquisar...")
    botao_pesquisar = driver.find_element(By.ID, "btnPesquisarFicha")
    botao_pesquisar.click()
    
    print(">>> Pesquisa enviada! Aguardando a tabela aparecer (Automático)...")
    try:
        # Configura um "vigia" que espera no MÁXIMO 30 segundos
        wait = WebDriverWait(driver, 30)
    
     
        wait.until(EC.presence_of_element_located((By.XPATH, "//tr[.//span[@title='WORKFLOW']]")))
    except Exception as e:
        print("ALERTA: A tabela demorou mais de 30s ou não trouxe resultados.")
   
    
    print(">>> Tabela detectada! Seguindo imediatamente...")
    
    # --- AQUI COMEÇA A FILTRAGEM AUTOMÁTICA DA TABELA ---
    print("\n" + "="*40)
    print("TABELA CARREGADA (ASSUMINDO). APLICANDO FILTROS DE COLUNA...")

    # 1. Filtrar Coluna 9 (supergedex)
    print(" -> Filtrando coluna 9 (supergedex)...")
    filtro_supergedex = driver.find_element(By.CSS_SELECTOR, "input[data-index='9']")
    filtro_supergedex.clear()
    filtro_supergedex.send_keys("supergedex")
    time.sleep(1) # Pequena pausa para garantir que o texto entrou
    filtro_supergedex.send_keys(Keys.ENTER) # O Enter confirma o filtro
    
    print("    [OK] Aguardando 1.5s para a tabela atualizar...")
    time.sleep(1.5) 

    # 2. Filtrar Coluna 5 (SE_ELTC)
    print(" -> Filtrando coluna 5 (SE_ELTC)...")
    filtro_eltc = driver.find_element(By.CSS_SELECTOR, "input[data-index='5']")
    filtro_eltc.clear()
    filtro_eltc.send_keys("SE_ELTC")
    time.sleep(1)
    filtro_eltc.send_keys(Keys.ENTER)
    
    print("    [OK] Aguardando 1.5s para a tabela atualizar...")
    time.sleep(1.5)

except Exception as e:
    print(f"\nERRO CRÍTICO: {e}")
    print("Dica: Se o erro for 'no such element', aumente o tempo de espera de 15s para 20s.")

# --- CHECKPOINT FINAL ---
print("\n" + "="*40)
print("SUCESSO: Filtros aplicados.")
print("Olhe para o Chrome: sobraram apenas os itens corretos?")
print("="*40 + "\n")

print("INICIANDO A EXTRAÇÃO DETALHADA (DRILL DOWN)...")
# Lista para guardar os dados

dados_coletados = []

try:
    print("\n" + "="*40)
    print("AGUARDANDO TABELA E APLICANDO FILTROS...")

    # --- ETAPA 1: GARANTIR QUE A TABELA CARREGOU ---
    wait = WebDriverWait(driver, 30)
    # Espera até aparecer pelo menos uma linha com botão de workflow
    wait.until(EC.presence_of_element_located((By.XPATH, "//tr[.//span[@title='WORKFLOW']]")))
    print(">>> Tabela detectada! Aplicando filtros...")
    time.sleep(2) # Pausa de segurança para o carregamento visual

    # --- ETAPA 2: APLICAR FILTROS (ANTES DE CONTAR AS LINHAS) ---
    
    # Filtro 1: supergedex (Coluna 9)
    try:
        f_supergedex = driver.find_element(By.CSS_SELECTOR, "input[data-index='9']")
        f_supergedex.clear()
        f_supergedex.send_keys("supergedex")
        time.sleep(0.5)
        f_supergedex.send_keys(Keys.ENTER)
        print(" -> Filtro 'supergedex' aplicado.")
        time.sleep(3) # Espera tabela atualizar
    except Exception as e:
        print(f" [AVISO] Falha ao filtrar supergedex: {e}")

    # Filtro 2: SE_ELTC (Coluna 5)
    try:
        f_eltc = driver.find_element(By.CSS_SELECTOR, "input[data-index='5']")
        f_eltc.clear()
        f_eltc.send_keys("SE_ELTC")
        time.sleep(0.5)
        f_eltc.send_keys(Keys.ENTER)
        print(" -> Filtro 'SE_ELTC' aplicado.")
        time.sleep(3) # Espera tabela atualizar
    except Exception as e:
        print(f" [AVISO] Falha ao filtrar SE_ELTC: {e}")

    # --- ETAPA 3: EXTRAÇÃO (AGORA COM A TABELA FILTRADA) ---
    print("\n" + "="*40)
    print("INICIANDO EXTRAÇÃO NOS RESULTADOS FILTRADOS...")

    # Localiza as linhas visíveis AGORA
    linhas_xpath = "//tr[.//span[@title='WORKFLOW']]"
    linhas = driver.find_elements(By.XPATH, linhas_xpath)
    total_encontrado = len(linhas)
    
    print(f">>> Encontrei {total_encontrado} projetos filtrados. INICIANDO...")

    for i in range(total_encontrado):
        print(f"--- Processando linha {i+1} de {total_encontrado} ---")

        try:
            # 1. Recaptura a linha (evita erro de elemento obsoleto)
            linha_atual = driver.find_elements(By.XPATH, linhas_xpath)[i]
            
            # 2. Localiza o botão
            botao_workflow = linha_atual.find_element(By.CSS_SELECTOR, "span[title='WORKFLOW']")
            modal_abriu = False
            
            # 3. Estratégia de Clique (Scroll + Pausa + Clique JS)
            # Rola a tela para o botão ficar no meio (Isso evita falhas de clique)
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_workflow)
            time.sleep(1.0) # Aumentado para 1s para o site "assentar" antes de clicar

            # Tentativa de Clique (Silenciosa)
            try:
                # Clique Principal
                driver.execute_script("arguments[0].click();", botao_workflow)
                
                # Verifica se abriu em 1 segundo
                try:
                    WebDriverWait(driver, 1.0).until(EC.visibility_of_element_located((By.ID, "modalWorkflowFichaCorpo")))
                    modal_abriu = True
                except:
                    # Se falhou (o site ignorou), clica de novo imediatamente (Reforço)
                    driver.execute_script("arguments[0].click();", botao_workflow)
                    # Agora espera o tempo normal
                    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "modalWorkflowFichaCorpo")))
                    modal_abriu = True
            except:
                pass # Se der erro técnico, modal_abriu continua False e o script avisa abaixo

            if not modal_abriu:
                print("   [ERRO CRÍTICO] O modal não abriu. Pulando linha.")
                continue

            # --- DENTRO DO MODAL ---
            time.sleep(1.5) # Pausa para os textos do modal carregarem
            elemento_modal = driver.find_element(By.ID, "modalWorkflowFichaCorpo")
            
            # A) Descrição (Do Título do Modal)
            descricao = "-"
            try:
                descricao = driver.find_element(By.ID, "modalWorkflowFichaTituloTitulo").text.strip()
            except: pass
            
            # B) Dados Básicos
            versao = elemento_modal.get_attribute("data-rev-ficha")
            codigo = elemento_modal.get_attribute("data-name-ficha")

            # C) Status (Regra: Concluído > Edição > Aprovação)
            status_final = "INDEFINIDO"
            detalhe_status = "-"
            try:
                # Checa canto superior direito
                rotulo_global = elemento_modal.find_element(By.CSS_SELECTOR, ".widget-toolbar .label-xlg").text.upper()
                
                if "CONCLUÍDO" in rotulo_global:
                    status_final = "APROVADO"
                    detalhe_status = "Finalizado"
                else:
                    # Checa histórico (item do topo)
                    etiquetas = elemento_modal.find_elements(By.CSS_SELECTOR, ".timeline-item .widget-toolbar .label-white")
                    if etiquetas:
                        texto_etiqueta = etiquetas[0].text.upper()
                        detalhe_status = texto_etiqueta
                        
                        if "EDIÇÃO" in texto_etiqueta: status_final = "REPROVADO"
                        elif "APROVAÇÃO" in texto_etiqueta: status_final = "EM ANÁLISE"
                        else: status_final = f"EM ANDAMENTO ({texto_etiqueta})"
                    else:
                        status_final = "EM ANDAMENTO (Sem histórico)"
            except: status_final = "ERRO LEITURA"

            # D) Datas
            data_inicio = "-"
            try: data_inicio = elemento_modal.find_element(By.XPATH, ".//small[contains(., 'iniciado em')]//b").text
            except: pass

            data_atualizacao = "-"
            try:
                datas = elemento_modal.find_elements(By.CSS_SELECTOR, ".timeline-label b")
                if datas: data_atualizacao = datas[0].text
            except: pass

            print(f"   -> {codigo} | {status_final} | {data_atualizacao}")

            # Salvar na memória
            dados_coletados.append({
                "Código": codigo,
                "Versão": versao,
                "Descrição": descricao,
                "Status": status_final,
                "Detalhe Etiqueta": detalhe_status,
                "Data Início": data_inicio,
                "Última Atualização": data_atualizacao
            })

            # FECHAR MODAL
            webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            
            # Espera o modal sumir (Crucial para não clicar errado na próxima)
            try:
                WebDriverWait(driver, 3).until(EC.invisibility_of_element_located((By.ID, "modalWorkflowFichaCorpo")))
            except:
                time.sleep(1) # Sleep de segurança se o wait falhar

        except Exception as e:
            print(f"   [ERRO NA LINHA {i+1}]: {e}")
            # Tenta fechar modal se der erro
            try: webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            except: pass

    # --- SALVAR NO EXCEL ---
    if dados_coletados:
        df = pd.DataFrame(dados_coletados)
        # Ordena as colunas
        cols = ["Código", "Versão", "Status", "Detalhe Etiqueta", "Data Início", "Última Atualização", "Descrição"]
        cols_finais = [c for c in cols if c in df.columns]
        df = df[cols_finais]
        
        df.to_excel("Relatorio_GEDEX_Final.xlsx", index=False)
        print("\nSUCESSO: Relatorio_GEDEX_Final.xlsx gerado!")
    else:
        print("\nNenhum dado foi coletado.")

except Exception as e:
    print("ERRO CRÍTICO NO SCRIPT:")
    traceback.print_exc()

input("Pressione ENTER para encerrar...")