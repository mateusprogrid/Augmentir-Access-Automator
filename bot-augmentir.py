import os
import shutil
import time
import pandas as pd

# --- Imports do Selenium ---
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, WebDriverException

# --- 1. CONFIGURAÇÕES ---
# (Idêntico ao V4)
PASTA_INPUT = "INPUT"
PASTA_ARQUIVADOS = "ARQUIVADOS"
ARQUIVO_MASTER = "Base de Dados da Solicitação de Acesso Sistemas de Manufatura.xlsx"
ABA_MASTER = "Augmentir" 
LINHA_CABECALHO_INPUT = 3 
NOME_COLUNA_USUARIO = "Nome completo do usuário"
AUGMENTIR_URL = "https://app.augmentir.com/#/configure/users"
AUGMENTIR_EMAIL = "mateus_melo@colpal.com"


# --- 2. MÓDULO DE PROCESSAMENTO DE ARQUIVOS ---
# (Idêntico ao V4, está funcionando perfeitamente)

def processar_planilhas_input():
    print(f"--- Módulo 1: Processando arquivos da pasta '{PASTA_INPUT}' ---")
    
    os.makedirs(PASTA_INPUT, exist_ok=True)
    os.makedirs(PASTA_ARQUIVADOS, exist_ok=True)

    arquivos_input = [f for f in os.listdir(PASTA_INPUT) if f.endswith('.xlsx')]
    
    if not arquivos_input:
        print("Nenhum arquivo .xlsx encontrado na pasta INPUT.")
        return []

    lista_total_nomes = []
    lista_dfs_para_anexar = []
    
    for arquivo in arquivos_input:
        caminho_arquivo = os.path.join(PASTA_INPUT, arquivo)
        try:
            df = pd.read_excel(caminho_arquivo, header=LINHA_CABECALHO_INPUT)

            if df.empty:
                print(f"Arquivo '{arquivo}' não contém dados (abaixo da linha 4). Pulando.")
            else:
                if NOME_COLUNA_USUARIO in df.columns:
                    nomes_novos = df[NOME_COLUNA_USUARIO].dropna().tolist()
                    if nomes_novos:
                        print(f"Encontrados {len(nomes_novos)} nomes no arquivo '{arquivo}'.")
                        lista_total_nomes.extend(nomes_novos)
                        lista_dfs_para_anexar.append(df)
                    else:
                        print(f"Arquivo '{arquivo}' processado, mas sem nomes na coluna '{NOME_COLUNA_USUARIO}'.")
                else:
                    print(f"ERRO: O arquivo '{arquivo}' não contém a coluna obrigatória '{NOME_COLUNA_USUARIO}'.")

            shutil.move(caminho_arquivo, os.path.join(PASTA_ARQUIVADOS, arquivo))
            print(f"Arquivo '{arquivo}' processado e movido para '{PASTA_ARQUIVADOS}'.")

        except Exception as e:
            print(f"ERRO ao processar o arquivo '{arquivo}': {e}")
            
    if lista_dfs_para_anexar:
        print(f"\nAtualizando base de dados local: '{ARQUIVO_MASTER}'...")
        try:
            df_novos_dados = pd.concat(lista_dfs_para_anexar, ignore_index=True)
            
            with pd.ExcelWriter(ARQUIVO_MASTER, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                writer.book = pd.ExcelFile(ARQUIVO_MASTER, engine='openpyxl').book
                start_row = writer.book[ABA_MASTER].max_row
                df_novos_dados.to_excel(writer, sheet_name=ABA_MASTER, startrow=start_row, index=False, header=False)
            
            print(f"Sucesso: {len(df_novos_dados)} novas linhas adicionadas à aba '{ABA_MASTER}'.")

        except FileNotFoundError:
            print(f"AVISO: Arquivo '{ARQUIVO_MASTER}' não encontrado.")
            print("Criando um novo arquivo com os dados atuais...")
            df_novos_dados.to_excel(ARQUIVO_MASTER, sheet_name=ABA_MASTER, index=False)
        except KeyError:
            print(f"ERRO CRÍTICO: A aba '{ABA_MASTER}' não existe no arquivo '{ARQUIVO_MASTER}'!")
            print("Nenhum dado foi salvo. Verifique o nome da aba.")
            return []
        except Exception as e:
            print(f"ERRO ao atualizar a base de dados: {e}")
            
    return lista_total_nomes

# --- 3. MÓDULO DE VERIFICAÇÃO WEB (AUGMENTIR) ---
# (ATUALIZADO V5 - Lógica de espera 'staleness_of')

def verificar_augmentir(nomes_para_verificar):
    print(f"\n--- Módulo 2: Verificando {len(nomes_para_verificar)} nomes no Augmentir ---")
    
    driver = None
    try:
        options = webdriver.ChromeOptions()
        options.add_experimental_option('excludeSwitches', ['enable-logging'])
        service = Service() 
        driver = webdriver.Chrome(service=service, options=options)
        
    except WebDriverException as e:
        print("\n==================== ERRO CRÍTICO ====================")
        if "session not created" in str(e):
            print("Falha ao iniciar o Chrome. Motivo provável:")
            print("O 'chromedriver.exe' não foi encontrado ou a versão está errada.")
        else:
            print(f"Erro inesperado do WebDriver: {e}")
        print("========================================================")
        return

    wait = WebDriverWait(driver, 10) # Espera padrão de 10s
    login_wait = WebDriverWait(driver, 60)

    # --- Localizadores (Idêntico V4) ---
    EMAIL_FIELD = (By.NAME, "username") 
    NEXT_BUTTON = (By.TAG_NAME, "button")
    SEARCH_BAR = (By.XPATH, "//input[contains(@placeholder, 'Search')]")
    RESULT_TEXT = (By.XPATH, "//*[contains(text(), 'Matched')]")
    NO_DATA_TEXT = (By.XPATH, "//*[contains(text(), 'No data')]")

    try:
        driver.get(AUGMENTIR_URL)

        # --- Lógica de Login ---
        try:
            search_bar = wait.until(EC.visibility_of_element_located(SEARCH_BAR))
            print("Login já estava ativo. Indo direto para a verificação.")
        
        except TimeoutException:
            print("Login não detectado. Iniciando processo de autenticação...")
            email_input = wait.until(EC.visibility_of_element_located(EMAIL_FIELD))
            email_input.send_keys(AUGMENTIR_EMAIL)
            driver.find_element(*NEXT_BUTTON).click()
            print("Login enviado. Aguardando autenticação (até 60s)...")
            search_bar = login_wait.until(EC.visibility_of_element_located(SEARCH_BAR))
            print("Autenticação concluída com sucesso.")

        # --- Loop de Verificação (LÓGICA V5) ---
        for nome in nomes_para_verificar:
            nome_limpo = str(nome).strip()
            if not nome_limpo: continue
                
            print(f"Verificando: {nome_limpo}...")
            
            # --- INÍCIO DA MUDANÇA V5 ---
            
            # 1. Captura o elemento de resultado ATUAL (antes da busca)
            try:
                elemento_resultado_anterior = driver.find_element(*RESULT_TEXT)
            except NoSuchElementException:
                elemento_resultado_anterior = None # Não existia, o que é ok

            # 2. Executa a busca
            search_bar.clear()
            search_bar.send_keys(nome_limpo)
            
            # 3. ESPERA INTELIGENTE: Espera o resultado ANTERIOR (se existia)
            #    desaparecer da tela (ficar "stale" ou obsoleto).
            if elemento_resultado_anterior:
                try:
                    # Pausa o script AQUI até que o "Matched..." antigo suma
                    wait.until(EC.staleness_of(elemento_resultado_anterior))
                except TimeoutException:
                    # Se não sumiu, a página pode estar travada
                    print(f"  -> AVISO: Resultado anterior para {nome_limpo} não desapareceu. O resultado pode ser instável.")
            
            # 4. Agora que o resultado antigo sumiu (ou não existia),
            #    esperamos o NOVO resultado aparecer
            try:
                # Espera pelo NOVO "Matched..."
                resultado_elemento = wait.until(EC.visibility_of_element_located(RESULT_TEXT))
                resultado = resultado_elemento.text
                
                # 5. Analisa o NOVO resultado
                if "Matched 0" in resultado:
                    # Espera 1 segundo para o "No data" estabilizar
                    time.sleep(1)
                    try:
                        # Confirma que "No data" está visível
                        driver.find_element(*NO_DATA_TEXT)
                        print(f"  -> Resultado: {nome_limpo} - Não possui acesso!")
                    except NoSuchElementException:
                        print(f"  -> Resultado: {nome_limpo} - Já possui acesso!!! (Caso: 'Matched 0' mas dados encontrados)")
                else:
                    # "Matched 1 user" (ou mais)
                    print(f"  -> Resultado: {nome_limpo} - Já possui acesso!!!")
            
            except TimeoutException:
                # Se nem o "Matched..." novo apareceu em 10s, o site travou
                print(f"  -> ERRO: Não foi possível obter resultado para {nome_limpo}. (Timeout da busca)")
            
            # --- FIM DA MUDANÇA V5 ---
            
            time.sleep(1) # Pausa curta entre as buscas

    except Exception as e:
        print(f"\nERRO CRÍTICO durante a automação web: {e}")
    
    finally:
        if driver:
            print("--- Verificação no Augmentir concluída ---")
            driver.quit()

# --- 4. EXECUÇÃO PRINCIPAL ---
# (Idêntico ao V4)

def main():
    print("=========================================")
    print("== Iniciando Robô de Acesso Augmentir ==")
    print("=========================================\n")
    
    try:
        nomes_a_verificar = processar_planilhas_input()
        
        nomes_unicos = list(set(nomes_a_verificar))
        
        if nomes_unicos:
            verificar_augmentir(nomes_unicos)
        else:
            print("\nNenhuma planilha nova ou nome encontrado. Encerrando sem abrir o navegador.")
            
    except Exception as e:
        print(f"\nERRO INESPERADO no fluxo principal: {e}")
    
    print("\n=========================================")
    print("== Robô finalizado. ==")
    print("=========================================")
    input("Pressione Enter para fechar o terminal...")

if __name__ == "__main__":
    main()