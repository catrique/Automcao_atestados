# 🧬 Relatório de Revisão de Código e Estabilização

Este documento apresenta uma revisão técnica detalhada dos pontos críticos encontrados na arquitetura do sistema, destacando os riscos de concorrência, vulnerabilidades de I/O de arquivos e as soluções implementadas para garantir a estabilidade das operações sob a gestão da Gerência de Governo Digital.

---

## 🔍 1. Interação com Hardware e Risco de Concorrência Física (PyAutoGUI)

* **Ponto Crítico:** No arquivo `auth_service.py`, a autenticação do proxy institucional depende do controle da tela via emulação de teclado e cliques (`pyautogui.click` e `pyautogui.write`).
* **Análise de Risco:** Como o script realiza cliques baseados nas coordenadas centrais da tela (`screen_width // 2`), qualquer interação do operador humano (como mover o mouse, trocar de janela ou clicar fora do Chrome) nos primeiros 3 segundos da inicialização desviará o foco. Isso fará com que o usuário e a senha do proxy sejam digitados em texto limpo no aplicativo que estiver em primeiro plano.
* **Solução e Mitigação:** Foi inserido um aviso visual severo e temporizado no console (`logger.warning`), instruindo o usuário a interromper qualquer interação com o teclado ou mouse durante a janela crítica de injeção de dados.

---

## ⚡ 2. Inicialização Estática vs. Cache de Credenciais (`update_data.py`)

* **Ponto Crítico:** No construtor original (`__init__`) da classe `DataUpdater`, as configurações de endpoints e URLs eram injetadas de forma estática via `get_config`.
* **Análise de Risco:** Se o operador acessasse a aba de configurações da interface gráfica (`gui.py`), atualizasse um endpoint ou corrigisse as credenciais do Betha e clicasse em "Salvar", as instâncias já criadas em memória continuavam usando as propriedades antigas (em cache) até que o programa fosse totalmente fechado e reaberto.
* **Recomendação de Melhoria:**Recomenda-se a conversão dessas chamadas para propriedades computadas (`@property`) ou a execução manual do método `reload_settings()` antes de cada ciclo de loop de sincronização. Isso garante que a memória volátil do script leia instantaneamente as alterações salvas no disco.

---

## 🗄️ 3. Risco de Corrupção por Falha de I/O (Gravação do JSON)

* **Ponto Crítico:** Rotinas padrão de escrita que utilizam `with open(SETTINGS_PATH, "w")` para salvar as credenciais diretamente no arquivo original.
* **Análise de Risco:** Se o computador sofrer uma queda repentina de energia, travamento do sistema operacional ou oscilação de rede no momento exato da escrita do arquivo `settings.json`, o arquivo pode ser corrompido, truncado ou zerado (0 KB), destruindo permanentemente os parâmetros do sistema.
* **Recomendação de Melhoria:** Implementar a estratégia de **Gravação Atômica** no módulo `loaders.py`. O sistema deve passar a escrever as atualizações primeiramente em um arquivo temporário separado (ex: `.settings.json.tmp`) e, somente após a conclusão bem-sucedida e fechamento do buffer em disco, utilizar o método `os.replace()` para substituir o arquivo original de forma instantânea, eliminando o risco de arquivos corrompidos por interrupções físicas.

## 🌐 4. Conexões Redundantes e Gerenciamento de Instâncias (Singletons)

* **Ponto Crítico:** Múltiplas redefinições e aberturas de escopo para comunicação com a API do Google Sheets (`gspread`).
* **Análise de Risco:** Instanciar a classe `SheetsService` repetidamente dentro de loops ou em diferentes eventos de botões faz com que o script realize novos handshakes de autenticação com o Google Cloud Platform, gerando lentidão e estourando rapidamente a cota de requisições por minuto da API v4.
* **Solução e Mitigação:** Criação e centralização do módulo `config_global.py`. Ele age como um barramento centralizado (*Singleton*), garantindo que a aplicação inicialize a conexão com a planilha uma única vez durante o carregamento do sistema. Todos os demais módulos utilizam obrigatoriamente a mesma instância global (`from config_global import sheets`).