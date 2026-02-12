# Automação de Atestados

## Descrição do Projeto

Este projeto Python foi desenvolvido para otimizar e automatizar o processo de gestão de atestados. Ele se integra de forma eficiente com diversas plataformas, como a API Betha, Google Sheets, SOC e CallMeBot, e oferece uma interface gráfica de usuário (GUI) intuitiva para facilitar a interação e o gerenciamento das operações.

## Funcionalidades Principais

O sistema oferece um conjunto robusto de funcionalidades para simplificar a rotina de gestão de atestados:

*   **Sincronização de Bases de Dados:** Permite a atualização e sincronização de informações de consulta (CID, CRM, etc.) entre a API Betha e planilhas Google Sheets.
*   **Geração e Importação de Relatórios:** Facilita o download de relatórios do sistema SOC e a subsequente importação de dados para planilhas, agilizando a análise e o processamento.
*   **Lançamento de Atestados no Betha:** Automatiza a leitura de dados de planilhas e o lançamento de atestados diretamente no sistema Betha, reduzindo erros manuais e otimizando o tempo.
*   **Gerenciamento de Credenciais:** Oferece uma interface segura para a configuração e atualização de credenciais de acesso para todas as APIs integradas (Betha, SOC, Proxy).
*   **Atualização de Token Betha:** Funcionalidade dedicada para a renovação do token de acesso da API Betha, garantindo a continuidade das operações.
*   **Notificações Automatizadas:** Integração com CallMeBot para o envio de notificações importantes, mantendo os usuários informados sobre o status das automações.
*   **Interface Gráfica Amigável (GUI):** Desenvolvida com `customtkinter`, proporciona uma experiência de usuário moderna e eficiente para o controle de todas as funcionalidades.

## Tecnologias Utilizadas

O projeto é construído sobre uma base tecnológica sólida, utilizando as seguintes ferramentas e bibliotecas:

*   **Python:** Linguagem de programação principal, conhecida por sua versatilidade e vasta gama de bibliotecas.
*   **customtkinter:** Um toolkit para Python que permite a criação de interfaces gráficas modernas e personalizáveis.
*   **requests:** Biblioteca HTTP para Python, utilizada para realizar requisições e interagir com as APIs externas (Betha, SOC, CallMeBot).
*   **pandas:** Poderosa biblioteca para manipulação e análise de dados, essencial para o tratamento de planilhas e relatórios.
*   **Google Sheets API:** Permite a leitura e escrita de dados em planilhas Google, facilitando a integração com fluxos de trabalho baseados em nuvem.
*   **API Betha:** Interface de programação para interação com o sistema Betha, utilizada para gestão de dados e atestados.
*   **API SOC:** Utilizada para o download programático de relatórios do sistema SOC.
*   **CallMeBot API:** Serviço para envio de mensagens via WhatsApp, integrado para notificações automatizadas.

## Estrutura do Projeto

A organização do projeto segue uma estrutura modular para facilitar a manutenção e escalabilidade:

```
Automcao_atestados/
├── .vscode/                 # Configurações do ambiente de desenvolvimento VS Code
├── config/                  # Módulos de configuração do projeto
│   ├── __init__.py          # Inicialização do módulo de configuração
│   ├── loaders.py           # Lógica para carregamento e gerenciamento de configurações
│   └── settings.json        # Arquivo de configurações e credenciais (não versionado)
├── repositories/            # Módulos para interação com fontes de dados e repositórios
│   └── update_data.py       # Lógica para sincronização e atualização de dados
├── services/                # Módulos que encapsulam a lógica de negócio e integração com APIs
│   ├── auth_service.py      # Serviço de autenticação para diversas plataformas
│   ├── betha_service.py     # Serviço de integração com a API Betha
│   ├── callmebot_service.py # Serviço para envio de mensagens via CallMeBot
│   ├── sheets_service.py    # Serviço de integração com Google Sheets
│   ├── soc_service.py       # Serviço de integração com a API SOC
│   ├── utils_service.py     # Funções utilitárias e de apoio geral
│   └── validation_service.py# Serviço para validação de dados de entrada
├── .gitignore               # Arquivo que especifica arquivos e diretórios a serem ignorados pelo Git
├── estrutura.py             # Definição de estruturas de dados ou classes auxiliares
├── gui.py                   # Implementação da interface gráfica do usuário (ponto de entrada principal)
├── main.py                  # Ponto de entrada alternativo/legado
└── README.md                # Documentação principal do projeto
```

## Como Usar

### Pré-requisitos

Para executar este projeto, você precisará ter o **Python 3.x** instalado em seu sistema operacional.

### Instalação

Siga os passos abaixo para configurar o ambiente e instalar as dependências:

1.  **Clone o repositório:**

    ```bash
    git clone https://github.com/caledcresst-dev/Automcao_atestados.git
    cd Automcao_atestados
    ```

2.  **Crie e ative um ambiente virtual (altamente recomendado):**

    ```bash
    python -m venv venv
    # No Linux/macOS:
    source venv/bin/activate
    # No Windows:
    .\venv\Scripts\activate
    ```

3.  **Instale as dependências:**

    ```bash
    pip install customtkinter requests pandas google-auth-oauthlib google-api-python-client
    ```

### Configuração

O projeto gerencia suas configurações através de um arquivo `settings.json` localizado no diretório `config/`. Este arquivo é essencial para armazenar credenciais e parâmetros de acesso às APIs.

Crie o arquivo `config/settings.json` seguindo o modelo abaixo:

```json
{
  "betha": {
    "api": {
      "base_url": "https://api",
      "url_login": "https://logi",
      "endpoints": {
        "atestado": "se",
        "cid": "",
        "medico": "?",
        "tipo_afastamento": "?",
        "tipo_atestado": "?",
        "motivo_consulta": "?",
        "pessoa_juridica": "//?",
        "listagem_matricula": "",
        "anexo": "/"
      },
      "user_access": "==",
      "authorization": "Bearer "
    },
    "user": {
      "admin": {
        "LOGIN": "SUA_CHAVE_CRIPTOGRAFADA",
        "PASSWORD": "SUA_SENHA_CRIPTOGRAFADA"
      }
    }
  },
  "soc": {
    "URL_SOC": "https://",
    "user": {
      "admin": {
        "LOGIN": "SEU_LOGIN_SOC",
        "PASSWORD": "SUA_SENHA_SOC",
        "SENHA_VIRTUAL": "SUA_SENHA_VIRTUAL"
      }
    }
  },
  "paths": {
    "downloads": "caminho/para/downloads",
    "logs": "caminho/para/logs"
  },
  "proxy": {
    "PROXY_HOST": "",
    "PROXY_PORT": "",
    "PROXY_USER": "",
    "PROXY_PASS": ""
  },
  "google_sheets": {
    "planilha": "nome da planilha",
    "aba": "nome da aba"
  }
}
```

> **Nota:** Os campos de login e senha no sistema utilizam chaves criptografadas. Certifique-se de configurar corretamente os endpoints e URLs de acordo com o seu ambiente de produção.

### Execução

Para iniciar a aplicação e sua interface gráfica, execute o arquivo `gui.py`:

```bash
python gui.py
```

## Interface Gráfica do Usuário (GUI)

A aplicação apresenta uma interface gráfica moderna e funcional, desenvolvida com `customtkinter`, que centraliza todas as operações de automação.
![alt text](<Captura de tela 2026-02-12 120847.png>)

### Seções da GUI:

*   **Menu Lateral Esquerdo:**
    *   **Atualizar Bases:** Sincroniza dados de CID, CRM e outros parâmetros.
    *   **Relatórios SOC:** Gerencia o download e importação de dados do SOC.
    *   **Lançar Atestados:** Lê as planilhas configuradas e realiza os lançamentos no sistema Betha.
    *   **Credenciais:** Abre o painel direito para configuração de logins.
*   **Painel de Credenciais (Direito):**
    *   Permite a inserção e salvamento de dados para Betha Cloud, Integração SOC e configurações de Proxy diretamente no `settings.json`.

## Contribuição

1.  Faça um fork do repositório.
2.  Crie uma nova branch (`git checkout -b feature/nova-feature`).
3.  Faça commit de suas alterações (`git commit -m 'Adiciona nova funcionalidade'`).
4.  Envie para a branch (`git push origin feature/nova-feature`).
5.  Abra um Pull Request.

## Licença

Este projeto está licenciado sob a licença MIT.

## Autor

Caled Cresst (caledcresst-dev)
