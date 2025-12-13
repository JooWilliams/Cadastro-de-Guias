import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time

# ==================== CONFIGURAÇÕES ====================
# Modifique estes valores conforme necessário
MES_REFERENCIA = "12 - dezembro"
STATUS_INICIAL = "EMITIDO"
DATA_VALIDADE = "31/12/2025"  # Formato brasileiro: DD/MM/AAAA
CAMINHO_PLANILHA = r"C:"  # Caminho para sua planilha
PORTA_DEBUG = 9222  # Porta do Chrome em modo debug
# =======================================================

# Definição dos grupos de especialidades
GRUPO_1 = ["TERAPIA OCUPACIONAL", "PSICOMOTRICIDADE", "MUSICOTERAPIA"]
GRUPO_2 = ["FONOAUDIOLOGIA", "PSICOPEDAGOGIA", "PSICOTERAPIA"]
GRUPO_3 = ["NUTRIÇÃO"]

# Mapeamento dos nomes da planilha para os nomes do sistema
MAPEAMENTO_ESPECIALIDADES = {
    # Variações possíveis -> Nome exato no sistema
    "terapia ocupacional": "TERAPIA OCUPACIONAL",
    "psicomotricidade": "PSICOMOTRICIDADE",
    "Musicoterapia": "MUSICOTERAPIA",
    "Fonoaudiologia": "FONOAUDIOLOGIA",
    "Psicopedagogia": "PSICOPEDAGOGIA",
    "Psicoterapia": "PSICOTERAPIA",
    "Nutrição": "NUTRIÇÃO"
}

def converter_nome_especialidade(especialidade_planilha):
    """Converte o nome da especialidade da planilha para o nome do sistema (MAIÚSCULAS)"""
    esp_normalizada = especialidade_planilha.strip().lower()
    
    # Busca no mapeamento
    if esp_normalizada in MAPEAMENTO_ESPECIALIDADES:
        return MAPEAMENTO_ESPECIALIDADES[esp_normalizada]
    
    # Se não encontrar no mapeamento, retorna em MAIÚSCULAS
    return especialidade_planilha.strip().upper()

def conectar_chrome_debug(porta=9222):
    """Conecta ao Chrome em modo debug"""
    options = webdriver.ChromeOptions()
    options.add_experimental_option("debuggerAddress", f"localhost:{porta}")
    driver = webdriver.Chrome(options=options)
    return driver

def normalizar_especialidade(especialidade):
    """Normaliza o nome da especialidade para comparação"""
    return especialidade.strip().lower()

def determinar_grupo(especialidade):
    """Determina a qual grupo a especialidade pertence"""
    esp_normalizada = normalizar_especialidade(especialidade)
    
    for esp in GRUPO_1:
        if normalizar_especialidade(esp) in esp_normalizada or esp_normalizada in normalizar_especialidade(esp):
            return 1
    
    for esp in GRUPO_2:
        if normalizar_especialidade(esp) in esp_normalizada or esp_normalizada in normalizar_especialidade(esp):
            return 2
    
    for esp in GRUPO_3:
        if normalizar_especialidade(esp) in esp_normalizada or esp_normalizada in normalizar_especialidade(esp):
            return 3
    
    return None

def agrupar_guias(df):
    """Agrupa as guias por beneficiário, senha e grupo de especialidades"""
    guias_agrupadas = []
    
    # Adiciona coluna de grupo
    df['Grupo'] = df['Especialidades'].apply(determinar_grupo)
    
    # Agrupa por Nome, Senha e Grupo
    grupos = df.groupby(['Nome do Beneficiário', 'Senha', 'Grupo'])
    
    for (nome, senha, grupo), grupo_df in grupos:
        especialidades = []
        for _, row in grupo_df.iterrows():
            especialidades.append({
                'nome': row['Especialidades'],
                'quantidade': row['Autorizada']
            })
        
        guias_agrupadas.append({
            'nome_beneficiario': nome,
            'senha': senha,
            'grupo': grupo,
            'especialidades': especialidades
        })
    
    return guias_agrupadas

def preencher_paciente(driver, wait, nome_paciente):
    """Preenche o campo de paciente"""
    campo_paciente = wait.until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/div[2]/form/div[1]/div[1]/div/div/div[1]/input"))
    )
    campo_paciente.clear()
    campo_paciente.send_keys(nome_paciente)
    time.sleep(1.5)  # Aguarda autocomplete carregar
    
    # Clica na primeira sugestão que aparece
    try:
        primeira_sugestao = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/main/div[2]/form/div[1]/div[1]/div/div/div[2]/div/span"))
        )
        primeira_sugestao.click()
        print(f"  ✓ Paciente selecionado: {nome_paciente}")
        return True
    except Exception as e:
        print(f"  ⚠ Paciente '{nome_paciente}' não encontrado no sistema")
        # Limpa o campo do paciente
        try:
            campo_paciente.clear()
        except:
            pass
        return False

def preencher_numero_guia(driver, wait, numero_guia):
    """Preenche o número da guia"""
    campo_guia = wait.until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/div[2]/form/div[2]/div[1]/input"))
    )
    campo_guia.clear()
    campo_guia.send_keys(str(numero_guia))

def adicionar_especialidade(driver, wait, especialidade, quantidade):
    """Adiciona uma especialidade com sua quantidade"""
    # Converte o nome da especialidade para o formato do sistema (MAIÚSCULAS)
    especialidade_sistema = converter_nome_especialidade(especialidade)
    
    # Seleciona a especialidade usando select_by_value (valores em MAIÚSCULAS)
    select_especialidade = Select(wait.until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/div[2]/form/div[3]/div[1]/div[1]/select"))
    ))
    
    try:
        select_especialidade.select_by_value(especialidade_sistema)
        print(f"    ✓ Especialidade selecionada: {especialidade_sistema}")
    except Exception as e:
        print(f"    ⚠ Erro: Especialidade '{especialidade_sistema}' não encontrada no sistema")
        print(f"    Original na planilha: '{especialidade}'")
        print(f"    Opções disponíveis:")
        for option in select_especialidade.options:
            if option.get_attribute('value'):
                print(f"      - {option.get_attribute('value')}")
        raise
    
    # Preenche a quantidade
    campo_quantidade = driver.find_element(By.XPATH, "/html/body/div[1]/main/div[2]/form/div[3]/div[1]/div[2]/input")
    campo_quantidade.clear()
    campo_quantidade.send_keys(str(quantidade))
    
    # Clica no botão para adicionar
    botao_adicionar = driver.find_element(By.XPATH, "/html/body/div[1]/main/div[2]/form/div[3]/div[1]/button")
    botao_adicionar.click()
    time.sleep(0.8)  # Aguarda a especialidade ser adicionada

def preencher_mes_referencia(driver, wait, mes_referencia):
    """Preenche o mês de referência"""
    select_mes = Select(wait.until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/div[2]/form/div[5]/div[1]/select"))
    ))
    select_mes.select_by_visible_text(mes_referencia)

def preencher_data_validade(driver, wait, data_validade):
    """Preenche a data de validade"""
    campo_data = wait.until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/div[2]/form/div[6]/input"))
    )
    campo_data.clear()
    campo_data.send_keys(data_validade)

def preencher_status_inicial(driver, wait, status_inicial):
    """Preenche o status inicial usando o valor do option"""
    select_status = Select(wait.until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/div[2]/form/div[7]/div/select"))
    ))
    # Usa select_by_value ao invés de select_by_visible_text
    select_status.select_by_value(status_inicial)

def cadastrar_guia(driver, wait, guia):
    """Cadastra uma guia completa no sistema"""
    try:
        print(f"\n{'='*60}")
        print(f"Cadastrando guia para: {guia['nome_beneficiario']}")
        print(f"Senha: {guia['senha']}")
        print(f"Grupo: {guia['grupo']}")
        print(f"Especialidades: {len(guia['especialidades'])}")
        
        # Clica no botão Nova Guia
        botao_nova_guia = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/main/div[1]/button"))
        )
        botao_nova_guia.click()
        time.sleep(1)
        
        # Preenche o paciente - se não encontrar, clica em voltar e pula esta guia
        if not preencher_paciente(driver, wait, guia['nome_beneficiario']):
            print("✗ Guia pulada: Paciente não encontrado")
            # Clica no botão Voltar do formulário de cadastro
            try:
                botao_voltar_form = wait.until(
                    EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/main/div[1]/div/button"))
                )
                botao_voltar_form.click()
                print("  ✓ Voltando para tela de listagem...")
                time.sleep(1)
            except Exception as e:
                print(f"  ⚠ Erro ao clicar em voltar: {str(e)}")
            return False
        
        preencher_numero_guia(driver, wait, guia['senha'])
        
        # Adiciona todas as especialidades
        for esp in guia['especialidades']:
            nome_convertido = converter_nome_especialidade(esp['nome'])
            print(f"  - Adicionando: {nome_convertido} (Qtd: {esp['quantidade']})")
            adicionar_especialidade(driver, wait, esp['nome'], esp['quantidade'])
        
        preencher_mes_referencia(driver, wait, MES_REFERENCIA)
        preencher_data_validade(driver, wait, DATA_VALIDADE)
        preencher_status_inicial(driver, wait, STATUS_INICIAL)
        
        # Clica no botão Criar Guia
        botao_criar = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/main/div[2]/form/div[8]/button[2]"))
        )
        botao_criar.click()
        
        print("✓ Guia cadastrada com sucesso!")
        time.sleep(2)  # Aguarda a página processar
        
        # Clica no botão Voltar para retornar à listagem de guias
        botao_voltar = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div/div[1]/button"))
        )
        botao_voltar.click()
        print("✓ Retornando para cadastrar próxima guia...")
        time.sleep(1.5)  # Aguarda voltar à tela inicial
        
        return True
        
    except Exception as e:
        print(f"✗ Erro ao cadastrar guia: {str(e)}")
        # Tenta voltar para a tela inicial em caso de erro
        try:
            # Primeiro tenta o botão voltar do formulário de cadastro
            try:
                botao_voltar_form = driver.find_element(By.XPATH, "/html/body/div[1]/main/div[1]/div/button")
                botao_voltar_form.click()
                print("  ✓ Voltando para tela de listagem (formulário)...")
                time.sleep(1)
            except:
                # Se não encontrar, tenta o botão voltar da página de detalhes
                try:
                    botao_voltar = driver.find_element(By.XPATH, "/html/body/div[1]/div/div[1]/div/div[1]/button")
                    botao_voltar.click()
                    print("  ✓ Voltando para tela de listagem (detalhes)...")
                    time.sleep(1)
                except:
                    print("  ⚠ Não foi possível voltar automaticamente")
        except:
            pass
        return False

def main():
    """Função principal"""
    print("="*60)
    print("SISTEMA DE CADASTRO AUTOMÁTICO DE GUIAS")
    print("="*60)
    print(f"\nConfigurações:")
    print(f"  Mês de Referência: {MES_REFERENCIA}")
    print(f"  Status Inicial: {STATUS_INICIAL}")
    print(f"  Data de Validade: {DATA_VALIDADE}")
    print(f"  Planilha: {CAMINHO_PLANILHA}")
    print("="*60)
    
    try:
        # Carrega a planilha
        print("\nCarregando planilha...")
        df = pd.read_excel(CAMINHO_PLANILHA)
        print(f"✓ Planilha carregada: {len(df)} registros encontrados")
        
        # Agrupa as guias
        print("\nAgrupando guias...")
        guias = agrupar_guias(df)
        print(f"✓ Total de guias a cadastrar: {len(guias)}")
        
        # Conecta ao Chrome
        print("\nConectando ao Chrome em modo debug...")
        driver = conectar_chrome_debug(PORTA_DEBUG)
        wait = WebDriverWait(driver, 10)
        print("✓ Conectado ao Chrome")
        
        # Cadastra as guias
        print("\nIniciando cadastro das guias...")
        sucesso = 0
        erro = 0
        puladas = 0
        
        for i, guia in enumerate(guias, 1):
            print(f"\n[{i}/{len(guias)}]", end=" ")
            resultado = cadastrar_guia(driver, wait, guia)
            
            if resultado == True:
                sucesso += 1
            elif resultado == False:
                # Verifica se foi erro ou guia pulada
                if "Paciente não encontrado" in str(resultado):
                    puladas += 1
                else:
                    erro += 1
            
            time.sleep(1)  # Pausa entre cadastros
        
        # Relatório final
        print("\n" + "="*60)
        print("RELATÓRIO FINAL")
        print("="*60)
        print(f"Total de guias processadas: {len(guias)}")
        print(f"✓ Cadastradas com sucesso: {sucesso}")
        print(f"⚠ Puladas (paciente não encontrado): {puladas}")
        print(f"✗ Erros: {erro}")
        print("="*60)
        
    except FileNotFoundError:
        print(f"\n✗ Erro: Arquivo '{CAMINHO_PLANILHA}' não encontrado!")
    except Exception as e:
        print(f"\n✗ Erro geral: {str(e)}")
    finally:
        print("\nProcesso finalizado. O navegador permanecerá aberto.")

if __name__ == "__main__":
    # Para iniciar o Chrome em modo debug, execute este comando no terminal:
    # Windows: chrome.exe --remote-debugging-port=9222 --user-data-dir="C:/selenium/chrome-profile"
    # Linux/Mac: google-chrome --remote-debugging-port=9222 --user-data-dir="/tmp/chrome-profile"
    
    main()