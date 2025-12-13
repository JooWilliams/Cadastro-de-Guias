# 🏥 Sistema de Cadastro Automático de Guias

Automação em Python + Selenium para cadastrar guias médicas automaticamente a partir de uma planilha Excel.

## 📋 Índice

- [Requisitos](#requisitos)
- [Instalação](#instalação)
- [Configuração](#configuração)
- [Preparação da Planilha](#preparação-da-planilha)
- [Como Usar](#como-usar)
- [Configurações Personalizáveis](#configurações-personalizáveis)
- [Agrupamento de Especialidades](#agrupamento-de-especialidades)
- [Solução de Problemas](#solução-de-problemas)

## 🔧 Requisitos

- Python 3.7 ou superior
- Google Chrome instalado
- Planilha Excel (.xlsx) com os dados das guias

## 📦 Instalação

### 1. Instalar as dependências Python

```bash
pip install pandas openpyxl selenium
```

### 2. Baixar o ChromeDriver (opcional)

O Selenium geralmente instala o driver automaticamente, mas se necessário, baixe em: https://chromedriver.chromium.org/

## ⚙️ Configuração

### 1. Iniciar o Chrome em Modo Debug

O programa precisa se conectar a uma sessão do Chrome já aberta. Execute um destes comandos:

**Windows:**
```bash
chrome.exe --remote-debugging-port=9222 --user-data-dir="C:/selenium/chrome-profile"
```

**Linux/Mac:**
```bash
google-chrome --remote-debugging-port=9222 --user-data-dir="/tmp/chrome-profile"
```

**Dica:** Salve este comando em um arquivo `.bat` (Windows) ou `.sh` (Linux/Mac) para facilitar.

### 2. Fazer login no sistema

Com o Chrome em modo debug aberto:
1. Acesse o sistema de guias
2. Faça login normalmente
3. Navegue até a página de listagem de guias
4. Deixe esta janela aberta

### 3. Configurar o script

Edite o arquivo `cadastrar_guias.py` e ajuste estas variáveis no topo do código:

```python
MES_REFERENCIA = "12 - Dezembro"  # Mês de referência das guias
STATUS_INICIAL = "EMITIDO"        # Status inicial das guias
DATA_VALIDADE = "31/12/2024"      # Data de validade (DD/MM/AAAA)
CAMINHO_PLANILHA = r"C:\caminho\para\sua\planilha.xlsx"
PORTA_DEBUG = 9222                # Porta do Chrome em debug
```

## 📊 Preparação da Planilha

A planilha Excel deve conter estas colunas **exatamente com estes nomes**:

| Nome do Beneficiário | Senha | Autorizada | Especialidades |
|---------------------|-------|------------|----------------|
| João da Silva       | 15032 | 5          | Terap Ocupac.  |
| João da Silva       | 15032 | 4          | Psicomotrici   |
| Maria Santos        | 23005 | 3          | Fonoaudiologia |

### Descrição das Colunas

- **Nome do Beneficiário**: Nome completo do paciente
- **Senha**: Número da guia (senha)
- **Autorizada**: Quantidade autorizada para a especialidade
- **Especialidades**: Nome da especialidade

### ⚠️ Importante

- Cada linha representa **uma especialidade** de uma guia
- Guias com mesmo **Nome** + **Senha** + **Grupo** serão cadastradas juntas
- Os nomes das especialidades podem estar em maiúsculas ou minúsculas (o sistema converte automaticamente)

## 🚀 Como Usar

### 1. Execute o script

```bash
python cadastrar_guias.py
```

### 2. Acompanhe o progresso

O programa mostrará em tempo real:
- Qual guia está sendo cadastrada
- Quais especialidades estão sendo adicionadas
- Status de sucesso ou erro
- Relatório final com estatísticas

### 3. Exemplo de saída

```
============================================================
SISTEMA DE CADASTRO AUTOMÁTICO DE GUIAS
============================================================

Configurações:
  Mês de Referência: 12 - Dezembro
  Status Inicial: EMITIDO
  Data de Validade: 31/12/2024
  Planilha: guias.xlsx
============================================================

Carregando planilha...
✓ Planilha carregada: 45 registros encontrados

Agrupando guias...
✓ Total de guias a cadastrar: 15

Conectando ao Chrome em modo debug...
✓ Conectado ao Chrome

Iniciando cadastro das guias...

[1/15] ============================================================
Cadastrando guia para: João da Silva
Senha: 15032
Grupo: 1
Especialidades: 2
  ✓ Paciente selecionado: João da Silva
  - Adicionando: Terapia ocupacional (Qtd: 5)
  - Adicionando: Psicomotricidade (Qtd: 4)
✓ Guia cadastrada com sucesso!
✓ Retornando para cadastrar próxima guia...

[2/15] ...
```

## 🎛️ Configurações Personalizáveis

### Mês de Referência

Formato: `"MM - Nome do Mês"`

Exemplos:
- `"01 - Janeiro"`
- `"06 - Junho"`
- `"12 - Dezembro"`

### Status Inicial

Valores aceitos:
- `"EMITIDO"` - Guia ou ficha foi emitida
- `"SUBIU"` - Subiu para análise
- `"ANALISE"` - Em processo de análise
- `"CANCELADO"` - Cancelado pelo sistema
- `"SAIU"` - Saiu da agenda
- `"RETORNOU"` - Retornou para a recepção
- `"NAO USOU"` - Não foi utilizado
- `"ASSINADO"` - Foi assinado completamente
- `"FATURADO"` - Processo de faturamento concluído
- `"ENVIADO A BM"` - Enviado para BM
- `"DEVOLVIDO BM"` - Devolvido pela BM
- `"PERDIDA"` - Guia perdida

### Data de Validade

Formato brasileiro: `DD/MM/AAAA`

Exemplo: `"31/12/2024"`

## 📚 Agrupamento de Especialidades

O sistema agrupa automaticamente as especialidades em 3 grupos. Guias com mesmo paciente, senha e grupo são cadastradas juntas.

### Grupo 1
- Terapia Ocupacional
- Psicomotricidade
- Musicoterapia

### Grupo 2
- Fonoaudiologia
- Psicopedagogia
- Psicoterapia

### Grupo 3
- Nutrição

### Exemplo Prático

**Planilha:**
| Nome           | Senha | Autorizada | Especialidades  |
|----------------|-------|------------|-----------------|
| João da Silva  | 15032 | 5          | Terap Ocupac.   |
| João da Silva  | 15032 | 4          | Psicomotrici    |
| João da Silva  | 23005 | 3          | Fonoaudiologia  |

**Resultado:**
- **Guia 1** (Senha 15032): Terapia Ocupacional (5) + Psicomotricidade (4)
- **Guia 2** (Senha 23005): Fonoaudiologia (3)

## 🔍 Especialidades Suportadas

O sistema converte automaticamente os nomes das especialidades:

| Na Planilha (flexível) | No Sistema (exato) |
|------------------------|-------------------|
| Terap Ocupac. / TERAPIA OCUPACIONAL | Terapia ocupacional |
| Psicomotrici / PSICOMOTRICIDADE | Psicomotricidade |
| MUSICOTERAPIA / Musicoterapia | Musicoterapia |
| FONOAUDIOLOGIA / Fonoaudiologia | Fonoaudiologia |
| PSICOPEDAGOGIA / Psicopedagogia | Psicopedagogia |
| PSICOTERAPIA / Psicoterapia | Psicoterapia |
| NUTRIÇÃO / Nutricao | Nutrição |
| FISIOTERAPIA | Fisioterapia |
| Avaliação neuropsicológica | Avaliação neuropsicológica |
| ARTETERAPIA | Arteterapia |
| Terapia ABA / TERAPIA ABA | Terapia ABA |

## 🐛 Solução de Problemas

### Erro: "Arquivo não encontrado"

**Causa:** Caminho da planilha incorreto

**Solução:** 
- Use `r` antes do caminho: `r"C:\Users\..."`
- Verifique se a extensão é `.xlsx` (não `.xlxs`)
- Confirme que o arquivo existe no caminho especificado

### Erro: "Could not locate element"

**Causa:** O Chrome em modo debug não está conectado ou a página não carregou

**Solução:**
1. Verifique se o Chrome está rodando na porta 9222
2. Confirme que você está logado no sistema
3. Navegue até a página de listagem de guias antes de executar o script
4. Aguarde a página carregar completamente

### Erro: "Especialidade não encontrada"

**Causa:** Nome da especialidade na planilha não está no mapeamento

**Solução:**
Adicione o mapeamento no dicionário `MAPEAMENTO_ESPECIALIDADES` do código:

```python
MAPEAMENTO_ESPECIALIDADES = {
    # ... mapeamentos existentes
    "seu nome na planilha": "Nome Exato no Sistema",
}
```

### O programa para no meio

**Possíveis causas:**
- Internet instável
- Página demorou para carregar
- Pop-up ou alerta inesperado

**Solução:**
- Execute novamente (ele continuará de onde parou se a planilha estiver ordenada)
- Aumente os tempos de espera (`time.sleep()`) no código
- Verifique se não há pop-ups bloqueando a tela

### Chrome fecha sozinho

**Causa:** Chrome não está em modo debug

**Solução:**
- Certifique-se de iniciar o Chrome com `--remote-debugging-port=9222`
- Não feche a janela do terminal que abriu o Chrome

## 📝 Notas Importantes

- ⚠️ **Não feche o navegador** durante a execução
- ⚠️ **Não clique na janela do Chrome** enquanto o script está rodando
- ⚠️ Mantenha a **janela visível** (não minimize)
- ✅ Teste primeiro com uma planilha pequena (5-10 guias)
- ✅ Faça backup da planilha original
- ✅ Verifique manualmente as primeiras guias cadastradas

## 📞 Suporte

Em caso de dúvidas ou problemas:

1. Verifique se todas as configurações estão corretas
2. Teste com uma planilha de exemplo pequena
3. Confirme que o Chrome está em modo debug
4. Verifique os logs de erro no console

## 📄 Licença e Direitos Autorais

© 2024 - Todos os direitos reservados.

Este software e sua documentação são propriedade exclusiva do autor. É proibida a reprodução, distribuição, modificação ou uso comercial sem autorização prévia por escrito.

**Uso permitido apenas para fins pessoais e internos da organização autorizada.**

---

**Desenvolvido com ❤️ usando Python + Selenium**