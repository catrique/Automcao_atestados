# 🔄 Máquina de Estados e Resiliência a Falhas

Este documento mapeia os estados lógicos da aplicação durante o ciclo de automação do CRESST. Ele serve para que a equipe de suporte ou futuros desenvolvedores da Gerência de Governo Digital compreendam as transições de fluxo e as políticas de tratamento de erro programadas.

---

## 📊 Diagrama de Estados do Sistema

O motor da automação transita entre estados rígidos de execução. Se um estado falhar, o sistema captura a exceção para evitar o fechamento inesperado (*crash*) da interface gráfica.

```text
  [ ESTADO: IDLE (Repouso) ]
               │
               ▼ (Gatilhador: Clique em Executar / Sincronizar)
  [ ESTADO: CARREGANDO_PARAMETROS ] ──► (Erro: JSON Corrompido) ──► [ ESTADO: PARADA_EMERGENCIA ]
               │
               ▼ (Sucesso)
  [ ESTADO: CONECTANDO_APIS ] ───────► (Falha de Rede / Timeout) ──► [ ESTADO: TRATAR_RETENTATIVA ]
               │                                                            │
               │◄─────────────────── (Ação: Tentar Novamente) ──────────────┘
               ▼
  [ ESTADO: EXTRAÇÃO_SOC_ATIVA ] ────► (Falha: Mudança de Layout Web) ──► [ ESTADO: INTERROMPER_ESTEIRA ]
               │
               ▼ (Download Concluído)
  [ ESTADO: PIPELINE_VALIDACAO ] ────► (Erro: Planilha Aberta/Bloqueada) ─► [ ESTADO: ALERTA_AO_OPERADOR ]
               │
               ▼ (Dados Higienizados)
  [ ESTADO: TRANSMISSÃO_BETHA ] ─────► (Erro 401: Token Expirado) ─────► [ ESTADO: REFRESH_TOKEN_AUTO ]
               │                                                                 │
               │◄─────────────────── (Token Renovado com Sucesso) ───────────────┘
               ▼ (Fim do Lote)
  [ ESTADO: GRAVAÇÃO_AUDITORIA ] ────► (Retorna ao Repouso) ──► [ ESTADO: IDLE ]

```

---

## 🔬 Detalhamento dos Estados Lógicos

### 1. `IDLE` (Aguardando Comando)

* **Descrição:** Estado de repouso inicial da interface CustomTkinter (`gui.py`). Não consome recursos de rede ou processamento pesado. O console exibe o histórico passado, aguardando uma nova interação do usuário.

### 2. `CARREGANDO_PARAMETROS`

* **Descrição:** O sistema aciona o módulo `loaders.py` para ler o arquivo `config/settings.json`.
* **Mecanismo de Resiliência:** Caso ocorra um erro de sintaxe ou o arquivo esteja ilegível, o sistema impede a inicialização de tarefas de rede, imprime uma mensagem contendo o erro técnico no console central e entra em `PARADA_EMERGENCIA` de forma controlada.

### 3. `CONECTANDO_APIS`

* **Descrição:** A aplicação testa os barramentos globais configurados em `config_global.py` e valida a comunicação inicial com as planilhas do Google.
* **Mecanismo de Resiliência:** Caso falte conexão com a internet ou o DNS não responda, o bloco `ErrorTranslator.traduzir()` intercepta a falha e avisa se o problema é de rede local ou instabilidade externa, disparando tentativas de reconexão de acordo com o método chamado.

### 4. `EXTRAÇÃO_SOC_ATIVA`

* **Descrição:** O Selenium assume o controle do navegador Chrome em segundo plano ou visível, navegando no painel do SOC para extrair os relatórios.
* **Mecanismo de Resiliência:** Se a sessão do SOC expirar no meio do download ou se o layout web falhar ao carregar, o robô dispara uma exceção customizada (`SessionExpired` ou `TimeoutException`), limpa a instância do driver da memória (`driver.quit()`) e atualiza o log avisando o operador que a esteira foi abortada para segurança dos dados.

### 5. `PIPELINE_VALIDACAO`

* **Descrição:** O Pandas assume o processamento da planilha local gerada na fase anterior, aplicando as higienizações de CPF, CRM e CIDs.
* **Mecanismo de Resiliência:** Se o arquivo Excel estiver bloqueado por estar aberto na tela do operador, o Python captura o erro de permissão (`PermissionError`) e emite uma janela de aviso na tela (`messagebox.showerror`), aguardando que o usuário feche o Excel para poder salvar as modificações com segurança.

### 6. `TRANSMISSÃO_BETHA`

* **Descrição:** Chamadas POST síncronas sequenciais enviando os atestados higienizados para os endpoints da API da Betha Cloud.
* **Mecanismo de Resiliência:** Se a API retornar um código `401 Unauthorized` indicando que o token expirou durante a transmissão de um lote grande, o sistema intercepta o erro, pausa o loop de envio,e aguarda a renovação do token.

### 7. `GRAVAÇÃO_AUDITORIA`

* **Descrição:** Fase final onde o status e as informações do operador (Nome, IP, Timestamp) coletados via `utils_service.py` são carimbados de volta no Google Sheets. Após o término, o sistema limpa os caches temporários de abas e retorna ao estado de repouso `IDLE`.
