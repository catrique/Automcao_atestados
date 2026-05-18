# 🗺️ Documentação de Handoff Técnico (Continuidade Operacional)

Este documento serve como um guia estratégico para a equipa técnica ou de suporte que assumirá a manutenção e a operação do sistema de automação do CRESST. O objetivo aqui é detalhar o comportamento de bastidores, as dependências de terceiros e os pontos de fragilidade conhecidos para garantir que o software sobreviva a longo prazo.

---

## 📌 Contexto e Propósito do Sistema

O sistema funciona como um *middleware* (barramento integrador) local. Ele resolve o atraso crónico no lançamento de licenças médicas no ERP municipal. 
Se a automação parar de funcionar, o impacto é imediato: o setor do CRESST volta a acumular pilhas de atestados, gerando retrabalho e potenciais erros de folha de pagamento por falta de lançamentos em tempo hábil.

---

## 🤝 Matriz de Responsabilidades e Resolução de Problemas

O sucesso da execução deste script depende da integração de quatro pilhas tecnológicas. Caso o software falhe ou apresente instabilidades, a **Gerência de Governo Digital** é a unidade central responsável por toda a infraestrutura, suporte, parametrização e articulação técnica com os serviços externos. Use esta tabela para isolar a origem do problema:

| Componente / Plataforma | Gestão e Suporte Integrado | Interface de Comunicação | Sintoma se Estiver Fora do Ar |
| :--- | :--- | :--- | :--- |
| **Plataforma SOC** | Gerência de Governo Digital | Automação de Interface UI (Selenium Web WebDriver) | O robô abre o Chrome, mas não consegue clicar nos botões ou o layout da plataforma mudou. |
| **Google Cloud Platform** | Gerência de Governo Digital | API v4 via biblioteca `gspread` | Logs na interface indicando erro de autenticação ou "Cota de requisições excedida". |
| **Gateway Betha Cloud** | Gerência de Governo Digital | Endpoints HTTP REST (JSON) | Erros de conexão recusada, Timeout (45s) ou erro 500 retornado pelo servidor da Betha. |
| **Firewall / Proxy Local** | Gerência de Governo Digital | Arquivo de Script `config_network.py` | Erros de Handshake SSL e falha de DNS ao tentar resolver as rotas das APIs externas. |

---

## ⚠️ Riscos Operacionais e Pontos Frágeis Conhecidos

Foram mapeados os pontos mais sensíveis da arquitetura atual que exigem atenção redobrada da nova equipa:

### 1. Injeção Física de Credenciais no Proxy (PyAutoGUI)
* **Onde está:** No módulo `auth_service.py`, dentro da rotina de autenticação do navegador Chrome.
* **O risco:** Para vencer o popup de autenticação do proxy da rede governamental, o script emula o teclado físico (`pyautogui.write`). Se o operador clicar em qualquer outra janela ou tirar o foco do navegador nos primeiros 3 segundos de execução, as senhas criptografadas serão injetadas em texto limpo no lugar errado.
* **Instrução de suporte:** Orientar os operadores a **nunca mexerem no mouse ou teclado** assim que clicarem no botão de início até que o Chrome estabilize.

### 2. Rigidez de Cabeçalhos na Planilha Excel
* **Onde está:** No validador local de arquivos (`validation_service.py`).
* **O risco:** O Pandas faz a leitura posicional e nominal de colunas exatas como `"Matrícula"`, `"Data Demissão"` e `"Nome"`. Se o fornecedor do SOC mudar uma letra nestes relatórios padronizados ou se um operador editar o arquivo manualmente, o validador quebrará com um erro de chave (`KeyError`).

### 3. Falta de Portabilidade do Arquivo de Configuração
* **Onde está:** No mecanismo de criptografia simétrica local.
* **O risco:** A semente geradora da chave Fernet utiliza `platform.node()` e `getpass.getuser()`. Isso significa que o arquivo `settings.json` configurado na Máquina do Operador A **está completamente ilegível** na Máquina do Operador B.
* **Instrução de suporte:** Se for migrar o sistema de computador, não copie o `settings.json`. Deixe o sistema gerar um novo e redigite as credenciais através do painel "⚙️ Credenciais" da interface gráfica na máquina de destino.

---

## 🛠️ Manutenção Recomendada para o Futuro Próximo

Para as próximas melhorias de estabilização da nova equipa, recomenda-se:
1. **Migração do Proxy:** Substituir o uso do PyAutoGUI por uma extensão auto-injetável de proxy criada em tempo de execução no Selenium, eliminando a dependência do foco na tela do Windows.
2. **Camada de Schema:** Implementar uma validação prévia de colunas com mensagens amigáveis em vez de deixar o Pandas disparar exceções nativas no console de logs.

---

## ✅ Checklist de Entrada para a Nova Equipa

Antes de modificar qualquer linha de código, garanta que cumpriu estes quatro passos:

- [ ] **Acesso ao Google Console:** Confirmar se tem o arquivo `credentials.json` original associado a uma conta de serviço com permissão de "Editor" na planilha master do Google Sheets.
- [ ] **Validação Prévia de Dados:** Como as integrações operam diretamente na base oficial da Betha, certifique-se de validar minuciosamente os arquivos gerados no validador local de Excel antes de acionar o envio de lotes para evitar lançamentos incorretos no sistema de produção.
- [ ] **Compilação Local:** Executar o PyInstaller localmente para garantir que o antivírus corporativo ou as políticas de grupo (GPO) do Windows não bloqueiam os binários temporários gerados na pasta `_MEIPASS`.
- [ ] **Rotação de Secrets:** Garantir que todas as senhas introduzidas na GUI foram salvas usando o botão "Salvar Configurações", validando se foram corretamente cifradas em disco.