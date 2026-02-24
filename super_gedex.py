from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By # <--- IMPORTANTE: Adicionei esta biblioteca
import time
import traceback
from selenium.webdriver.common.keys import Keys
import pandas as pd # Para criar o Excel
import re # Para achar datas no texto
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# --- SEUS DADOS ---
# Preencha aqui para não ter que digitar toda vez
MINHA_MATRICULA = "Login"
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
    
    print(">>> Pesquisa enviada! O robô vai esperar 25 segundos pela tabela...")
    time.sleep(25) # Tempo para o GEDEX carregar os resultados
    
    # --- AQUI COMEÇA A FILTRAGEM AUTOMÁTICA DA TABELA ---
    print("\n" + "="*40)
    print("TABELA CARREGADA (ASSUMINDO). APLICANDO FILTROS DE COLUNA...")

    # 1. Filtrar Coluna 9 (filtro)
    print(" -> Filtrando coluna 9 (filtro)...")
    filtro_coluna9 = driver.find_element(By.CSS_SELECTOR, "input[data-index='9']")
    filtro_coluna9.clear()
    filtro_coluna9.send_keys("filtro")
    time.sleep(1) # Pequena pausa para garantir que o texto entrou
    filtro_coluna9.send_keys(Keys.ENTER) # O Enter confirma o filtro
    
    print("    [OK] Aguardando 3s para a tabela atualizar...")
    time.sleep(3) 

    # 2. Filtrar Coluna 5 (SE_ELTC)
    print(" -> Filtrando coluna 5 (SE_ELTC)...")
    filtro_eltc = driver.find_element(By.CSS_SELECTOR, "input[data-index='5']")
    filtro_eltc.clear()
    filtro_eltc.send_keys("SE_ELTC")
    time.sleep(1)
    filtro_eltc.send_keys(Keys.ENTER)
    
    print("    [OK] Aguardando 3s para a tabela atualizar...")
    time.sleep(3)

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
    # Localiza as linhas que têm Workflow
    linhas_xpath = "//tr[.//span[@title='WORKFLOW']]"
    wait_geral = WebDriverWait(driver, 15)
    wait_geral.until(EC.presence_of_element_located((By.XPATH, linhas_xpath)))
    
    linhas = driver.find_elements(By.XPATH, linhas_xpath)
    total_encontrado = len(linhas)
    
# --- MODO COMPLETO (SEM LIMITE) ---
    # Agora o loop vai de 0 até o total de linhas encontradas na tela
    print(f">>> Encontrei {total_encontrado} projetos. INICIANDO EXTRAÇÃO COMPLETA...")

    for i in range(total_encontrado):
        

        try:
            # 1. RECAPTURA A LINHA (Anti-erro)
            linha_atual = driver.find_elements(By.XPATH, linhas_xpath)[i]
            
            # 2. PEGA DESCRIÇÃO
            try:
                descricao = linha_atual.find_element(By.CSS_SELECTOR, "td.hidden-800").text.strip()
            except:
                descricao = "-"

            # 3. CLIQUE ROBUSTO (Scroll + Retry)
            botao_workflow = linha_atual.find_element(By.CSS_SELECTOR, "span[title='WORKFLOW']")
            modal_abriu = False
            
            for tentativa in range(3):
                try:
                    # Centraliza o botão na tela
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", botao_workflow)
                    time.sleep(1)
                    # Clica via Javascript
                    driver.execute_script("arguments[0].click();", botao_workflow)
                    # Espera o modal
                    WebDriverWait(driver, 5).until(EC.visibility_of_element_located((By.ID, "modalWorkflowFichaCorpo")))
                    modal_abriu = True
                    break 
                except:
                    print(f"   (Tentativa de clique {tentativa+1} falhou...)")
            
            if not modal_abriu:
                print("   [ERRO] Não consegui abrir o workflow. Pulando.")
                continue

            # --- 4. DENTRO DO MODAL ---
            # Pequena pausa para garantir que os textos carregaram dentro do modal
            time.sleep(1.5)
            elemento_modal = driver.find_element(By.ID, "modalWorkflowFichaCorpo")
            
            # --- NOVA CAPTURA DE DESCRIÇÃO (ID: modalWorkflowFichaTituloTitulo) ---
            descricao = "-"
            try:
                # Pega o texto do H6 que contém o título do painel
                descricao = driver.find_element(By.ID, "modalWorkflowFichaTituloTitulo").text.strip()
            except:
                pass
            # -----------------------------------------------------------------------
            
            # DADOS BÁSICOS
            versao = elemento_modal.get_attribute("data-rev-ficha")
            codigo = elemento_modal.get_attribute("data-name-ficha")

            # --- A NOVA LÓGICA DE STATUS ---
            status_final = "INDEFINIDO"
            detalhe_status = "-"
            
            try:
                # PASSO 1: Checar o Rótulo Global (Canto Superior Direito)
                # Procura a etiqueta grande (label-xlg)
                rotulo_global = elemento_modal.find_element(By.CSS_SELECTOR, ".widget-toolbar .label-xlg").text.upper()
                
                if "CONCLUÍDO" in rotulo_global:
                    status_final = "APROVADO"
                    detalhe_status = "Finalizado"
                
                else:
                    # PASSO 2: Se não está concluído, olha a ÚLTIMA CAIXINHA (Topo da timeline)
                    # Procura todas as etiquetas brancas (label-white) dentro dos itens da timeline
                    # Como a timeline é decrescente (mais recente no topo), pegamos o item [0]
                    etiquetas_timeline = elemento_modal.find_elements(By.CSS_SELECTOR, ".timeline-item .widget-toolbar .label-white")
                    
                    if etiquetas_timeline:
                        texto_etiqueta = etiquetas_timeline[0].text.upper()
                        detalhe_status = texto_etiqueta # Guarda o que leu para conferência
                        
                        if "EDIÇÃO" in texto_etiqueta:
                            status_final = "REPROVADO"
                        elif "APROVAÇÃO" in texto_etiqueta:
                            status_final = "EM ANÁLISE"
                        else:
                            status_final = f"EM ANDAMENTO ({texto_etiqueta})"
                    else:
                        status_final = "EM ANDAMENTO (Sem passos)"

            except Exception as e_status:
                print(f"   [Erro na lógica de status]: {e_status}")
                status_final = "ERRO LEITURA"

            # DATAS (Início e Atualização)
            data_inicio = "N/D"
            try:
                data_inicio = elemento_modal.find_element(By.XPATH, ".//small[contains(., 'iniciado em')]//b").text
            except: pass

            data_atualizacao = "N/D"
            try:
                # Pega a primeira data que aparece na timeline (a mais recente)
                datas = elemento_modal.find_elements(By.CSS_SELECTOR, ".timeline-label b")
                if datas:
                    data_atualizacao = datas[0].text
            except: pass

            print(f"   -> {codigo} | {status_final} (Baseado em: {detalhe_status})")

            # SALVAR
            dados_coletados.append({
                "Código": codigo,
                "Versão": versao,
                "Descrição": descricao,
                "Status": status_final,
                "Detalhe Etiqueta": detalhe_status,
                "Data Início": data_inicio,
                "Última Atualização": data_atualizacao
            })

            # 5. FECHAR MODAL
            webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            time.sleep(1.5)

        except Exception as e:
            print(f"   [ERRO GERAL LINHA {i+1}]: {e}")
            try: webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
            except: pass

    # EXPORTAR
    if dados_coletados:
        df = pd.DataFrame(dados_coletados)
        # Reordenando colunas para facilitar leitura
        cols = ["Código", "Versão", "Status", "Detalhe Etiqueta", "Data Início", "Última Atualização", "Descrição"]
        # Filtra só as colunas que realmente existem
        cols_finais = [c for c in cols if c in df.columns]
        df = df[cols_finais]
        
        df.to_excel("Relatorio_GEDEX_Status_Corrigido.xlsx", index=False)
        print("\nSUCESSO: Relatorio_GEDEX_Status_Corrigido.xlsx gerado!")
    else:
        print("\nNenhum dado coletado.")

except Exception as e:
    print("ERRO CRÍTICO:")
    traceback.print_exc()

input("Pressione ENTER para encerrar...")