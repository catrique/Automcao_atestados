
# 📐 Fluxo Operacional Detalhado (Pipeline de Integração)

Este documento descreve as etapas lógicas e sequenciais que o sistema executa para extrair, sanitizar, validar e transmitir as licenças médicas do CRESST. Ele serve como um mapa de depuração para entender como os módulos se comunicam em tempo de execução.

---

## 🧭 Visão Geral do Circuito de Dados

O ciclo completo da automação é dividido em três grandes fases operacionais interdependentes:

```text
[ FASE 1: Extração (SOC) ] ──► [ FASE 2: Higienização (Pandas) ] ──► [ FASE 3: Carga (Betha API) ]

```

---

## 📂 Detalhamento das Fases

### 🏁 FASE 1: Captação e Extração de Relatórios (`soc_service.py`)

1. **Disparo por Thread:** O operador seleciona o intervalo de datas na interface (`gui.py`) e clica em iniciar. A interface gera uma Thread separada para evitar o congelamento da janela do Windows.
2. **Inicialização do WebDriver:** O Selenium instancia o Google Chrome utilizando as configurações de rede do `config_network.py` (forçando o desvio de proxies locais se necessário).
3. **Autenticação e Injeção:** O robô realiza o login na plataforma web do SOC. Caso haja popups institucionais de proxy, o script faz a temporização para a inserção das credenciais de rede.
4. **Download do Lote:** O robô navega até a tela de relatórios de afastamento, preenche o formulário com o período selecionado e solicita a exportação. Ele monitora a pasta de downloads até que o arquivo `.zip` seja completamente baixado, realizando a extração do arquivo Excel bruto para o diretório local de trabalho.

### ⚙️ FASE 2: Esteira de Higienização e Regras Clínicas (`validation_service.py`)

O arquivo bruto extraído do SOC não é enviado diretamente para o governo. Ele é carregado na memória via Pandas e passa por filtros estritos:

* **Normalização Nominal:** O método `normalizar_texto()` limpa acentuações estranhas, converte caracteres especiais (ex: `Ç` vira `C`) e remove de forma irreversível prefixos honoríficos (`DR.`, `DRA.`, `DR `, `DRA `) para evitar divergências na busca de nomes.
* **Sanitização de CPFs:** Aplica-se uma expressão regular (`re.sub(r'[^0-9]', '')`) que elimina pontos e traços, unificando todos os registros em strings limpas de 11 dígitos numéricos.
* **Validação contra Tabelas de Referência:**
* **Médicos:** O CRM e o nome do médico são cruzados com o mapa de referência carregado a partir do Google Sheets.
* **CIDs:** O código do CID é convertido para letras minúsculas e limpo de caracteres especiais (ex: `M54.5` ou `M54-5` viram `m545`) e validado contra a lista oficial de CIDs aceitos na base da Betha.
* **Vínculos:** O sistema verifica se a matrícula informada bate com o CPF do servidor ativo.


* **Separação de Erros:** As linhas que violarem qualquer uma dessas regras recebem uma descrição detalhada em uma nova coluna chamada `ERROS` diretamente no arquivo Excel local, isolando os registros corrompidos e impedindo que payloads inválidos cheguem à API.

### 🚀 FASE 3: Processamento de Matrículas e Carga REST (`betha_service.py`)

Os registros considerados 100% válidos na Fase 2 entram na fila de transmissão:

1. **Busca Preventiva de ID:** O sistema realiza uma chamada GET para o endpoint `buscar_matricula()` utilizando o CPF sanitizado do servidor. Isso é obrigatório para coletar o identificador interno único do funcionário dentro do ecossistema Betha.
2. **Montagem do Payload:** Com o ID do servidor em mãos, o script monta o JSON estruturado contendo a data de início da licença, dias de afastamento, o código do CID validado e o registro profissional do médico.
3. **Envio e Auditoria de Sucesso:** A requisição é disparada via método `POST` HTTP.
* Se o servidor da Betha responder com sucesso (`200 OK`), a automação aciona o módulo `sheets_service.py` para atualizar a linha equivalente no Google Sheets.
* A linha é carimbada com o status de sucesso, o nome do operador do sistema, o número de IP da máquina que realizou o envio e o timestamp (data/hora exata), criando uma trilha de auditoria para o Governo Digital.
