import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
import time
import datetime

# ==================== CONFIGURAÇÕES ====================
# Modifique estes valores conforme necessário
MES_REFERENCIA = "03 - março"
ANO_REFERENCIA = "2026"  # Ano de referência
PERIODICIDADE = "MES_FECHADO"  # NOVO: Opções: MES_FECHADO, SEMANAL, QUINZENAL
STATUS_INICIAL = "EMITIDO"
DATA_VALIDADE = "24/05/2026"  # Formato brasileiro: DD/MM/AAAA
CAMINHO_PLANILHA = r"C:\Users\vinic\OneDrive\Ambiente de Trabalho\guias-porto-extraidas.xlsx"  # Caminho para sua planilha
PORTA_DEBUG = 9223  # Porta do Chrome em modo debug
# =======================================================

# Definição dos grupos de especialidades
GRUPO_1 = ["TERAPIA OCUPACIONAL", "PSICOMOTRICIDADE", "MUSICOTERAPIA"]
GRUPO_2 = ["FONOAUDIOLOGIA", "PSICOPEDAGOGIA", "PSICOTERAPIA"]
GRUPO_3 = ["NUTRIÇÃO"]
GRUPO_4 = ["TERAPIA ABA"]

# Mapeamento dos nomes da planilha para os nomes do sistema
MAPEAMENTO_ESPECIALIDADES = {
    # Variações possíveis -> Nome exato no sistema
    "terapia ocupacional": "TERAPIA OCUPACIONAL",
    "psicomotricidade": "PSICOMOTRICIDADE",
    "Musicoterapia": "MUSICOTERAPIA",
    "Fonoaudiologia": "FONOAUDIOLOGIA",
    "Psicopedagogia": "PSICOPEDAGOGIA",
    "Psicoterapia": "PSICOTERAPIA",
    "Nutrição": "NUTRIÇÃO",
    "terapia aba": "TERAPIA ABA"
}

def converter_nome_especialidade(especialidade_planilha):
    """Converte o nome da especialidade da planilha para o nome do sistema (MAIÚSCULAS)"""
    esp_normalizada = especialidade_planilha.strip().lower()
    
    # Busca no mapeamento
    if esp_normalizada in MAPEAMENTO_ESPECIALIDADES:
        return MAPEAMENTO_ESPECIALIDADES[esp_normalizada]
    
    # Se não encontrar no mapeamento, retorna em MAIÚSCULAS
    return especialidade_planilha.strip().upper()

def conectar_chrome_debug(porta=9223):
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

    for esp in GRUPO_4:
        if normalizar_especialidade(esp) in esp_normalizada or esp_normalizada in normalizar_especialidade(esp):
            return 4
    
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
    time.sleep(1.3)  # Aguarda autocomplete carregar
    
    # Clica na primeira sugestão que aparece
    try:
        primeira_sugestao = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/main/div[2]/form/div[1]/div[1]/div/div/div[2]/div/span"))
        )
        primeira_sugestao.click()
        print(f"  ✓ Paciente: {nome_paciente}")
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
        print(f"    ✓ Especialidade: {especialidade_sistema}")
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
    time.sleep(0.6)  # Aguarda a especialidade ser adicionada

def preencher_mes_referencia(driver, wait, mes_referencia):
    """Preenche o mês de referência"""
    select_mes = Select(wait.until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/div[2]/form/div[5]/div[1]/select"))
    ))
    select_mes.select_by_visible_text(mes_referencia)

def preencher_ano_referencia(driver, wait, ano_referencia):
    """Preenche o ano de referência"""
    select_ano = Select(wait.until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/div[2]/form/div[5]/div[2]/select"))
    ))
    select_ano.select_by_value(str(ano_referencia))
    print(f"  ✓ Ano selecionado: {ano_referencia}")

def preencher_data_validade(driver, wait, data_validade):
    """Preenche a data de validade"""
    campo_data = wait.until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/div[2]/form/div[6]/input"))
    )
    campo_data.clear()
    campo_data.send_keys(data_validade)

def preencher_periodicidade(driver, wait, periodicidade):
    """Preenche a periodicidade (NOVO CAMPO)"""
    select_periodicidade = Select(wait.until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/div[2]/form/div[7]/div[1]/select"))
    ))
    select_periodicidade.select_by_value(periodicidade)
    
    # Mapeia os valores para nomes amigáveis para o print
    nomes_periodicidade = {
        "MES_FECHADO": "Mês fechado",
        "SEMANAL": "Semanal",
        "QUINZENAL": "Quinzenal"
    }
    nome_exibicao = nomes_periodicidade.get(periodicidade, periodicidade)
    print(f"  ✓ Período: {nome_exibicao}")

def preencher_status_inicial(driver, wait, status_inicial):
    """Preenche o status inicial usando o valor do option"""
    select_status = Select(wait.until(
        EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/div[2]/form/div[8]/div/select"))
    ))
    # Usa select_by_value ao invés de select_by_visible_text
    select_status.select_by_value(status_inicial)
    print(f"  ✓ Status: {status_inicial}")

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
        time.sleep(2.2)
        
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
            return False, "Paciente não encontrado"
        
        preencher_numero_guia(driver, wait, guia['senha'])
        
        # Adiciona todas as especialidades
        for esp in guia['especialidades']:
            nome_convertido = converter_nome_especialidade(esp['nome'])
            print(f"  - Adicionando: {nome_convertido} (Qtd: {esp['quantidade']})")
            adicionar_especialidade(driver, wait, esp['nome'], esp['quantidade'])
        
        preencher_mes_referencia(driver, wait, MES_REFERENCIA)
        preencher_ano_referencia(driver, wait, ANO_REFERENCIA)
        preencher_data_validade(driver, wait, DATA_VALIDADE)
        preencher_periodicidade(driver, wait, PERIODICIDADE)  # NOVO: Preenche a periodicidade
        preencher_status_inicial(driver, wait, STATUS_INICIAL)
        
        # Clica no botão Criar Guia
        botao_criar = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/main/div[2]/form/div[9]/button[2]"))
        )
        botao_criar.click()
        
        print("✓ Guia cadastrada com sucesso!")
        time.sleep(1.5)  # Aguarda a página processar
        
        # Clica no botão Voltar para retornar à listagem de guias
        botao_voltar = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div[1]/div/div[1]/div/div[1]/button"))
        )
        botao_voltar.click()
        time.sleep(2)  # Aguarda voltar à tela inicial
        
        return True, "Sucesso"
        
    except Exception as e:
        msg_erro = str(e)
        print(f"✗ Erro ao cadastrar guia: {msg_erro}")
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
        return False, f"Erro: {msg_erro}"

def main():
    """Função principal"""
    print("="*60)
    print("SISTEMA DE CADASTRO AUTOMÁTICO DE GUIAS")
    print("="*60)
    print(f"\nConfigurações:")
    print(f"  Mês de Referência: {MES_REFERENCIA}")
    print(f"  Ano de Referência: {ANO_REFERENCIA}")
    print(f"  Periodicidade: {PERIODICIDADE}")
    print(f"  Status Inicial: {STATUS_INICIAL}")
    print(f"  Data de Validade: {DATA_VALIDADE}")
    print(f"  Planilha: {CAMINHO_PLANILHA}")
    print("="*60)
    
    try:
        # Carrega a planilha da aba correta
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
        
        # Lista para armazenar guias com erro
        guias_com_erro = []
        
        for i, guia in enumerate(guias, 1):
            print(f"\n[{i}/{len(guias)}]", end=" ")
            resultado, mensagem = cadastrar_guia(driver, wait, guia)
            
            if resultado:
                sucesso += 1
            else:
                # Armazena os dados da guia e o motivo do erro
                guia_erro = {
                    "Beneficiário": guia['nome_beneficiario'],
                    "Senha": guia['senha'],
                    "Motivo": mensagem,
                    "Hora": datetime.datetime.now().strftime("%H:%M:%S")
                }
                guias_com_erro.append(guia_erro)
                
                if "Paciente não encontrado" in mensagem:
                    puladas += 1
                else:
                    erro += 1
            
            time.sleep(1)  # Pausa entre cadastros
        
        # Relatório final
        print("\n" + "="*60)
        print("RELATÓRIO FINAL")
        print(f"Guias processadas: {len(guias)}")
        print(f"✓ Sucesso: {sucesso}")
        print(f"⚠ Paciente não encontrado: {puladas}")
        print(f"✗ Erros: {erro}")
        print("="*60)
        
        # Gera arquivo de log se houver erros
        if guias_com_erro:
            timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
            nome_arquivo_erro = f"erros_guias_{timestamp}.txt"
            
            print(f"\nGerando log de erros: {nome_arquivo_erro}")
            try:
                with open(nome_arquivo_erro, "w", encoding="utf-8") as f:
                    f.write("="*60 + "\n")
                    f.write(f"RELATÓRIO DE ERROS DE CADASTRO - {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
                    f.write("="*60 + "\n\n")
                    
                    for item in guias_com_erro:
                        f.write(f"Beneficiário: {item['Beneficiário']}\n")
                        f.write(f"Senha: {item['Senha']}\n")
                        f.write(f"Motivo: {item['Motivo']}\n")
                        f.write(f"Hora: {item['Hora']}\n")
                        f.write("-" * 40 + "\n")
                
                print(f"Gerado '{nome_arquivo_erro}', Consulte para ver detalhes de erros")
            except Exception as e:
                print(f"✗ Erro ao criar arquivo de log: {e}")

    except FileNotFoundError:
        print(f"\n✗ Erro: Arquivo '{CAMINHO_PLANILHA}' não encontrado!")
    except Exception as e:
        print(f"\n✗ Erro geral: {str(e)}")
    finally:
        print("\nProcesso finalizado. O navegador permanecerá aberto.")

if __name__ == "__main__":
    # Para iniciar o Chrome em modo debug, execute este comando no terminal:
    # Windows: chrome.exe --remote-debugging-port=9223 --user-data-dir="C:/selenium/chrome-profile"
    # Linux/Mac: google-chrome --remote-debugging-port=9223 --user-data-dir="/tmp/chrome-profile"
    
    main()