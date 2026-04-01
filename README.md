# 🏥 Automação de Atestados — CRESST

> Sistema Python para automação completa do ciclo de gestão de atestados médicos, com integração entre Betha Cloud, SOC, Google Sheets e notificações via WhatsApp.

---

## 📋 Índice

- [Sobre o Projeto](#sobre-o-projeto)
- [Funcionalidades](#funcionalidades)
- [Tecnologias](#tecnologias)
- [Estrutura do Projeto](#estrutura-do-projeto)
- [Pré-requisitos](#pré-requisitos)
- [Instalação](#instalação)
- [Configuração](#configuração)
- [Como Usar](#como-usar)
- [Interface Gráfica](#interface-gráfica)
- [Geração de Executável](#geração-de-executável)
- [Contribuição](#contribuição)
- [Licença](#licença)

---

## Sobre o Projeto

O **Automação de Atestados** nasceu da necessidade de eliminar o trabalho manual e repetitivo na gestão de atestados médicos. O sistema integra quatro plataformas — **Betha Cloud**, **SOC**, **Google Sheets** e **CallMeBot** — em um único fluxo automatizado, controlado por uma interface gráfica moderna desenvolvida com `customtkinter`.

O fluxo principal funciona assim:

```
SOC (relatório) → Validação Excel → Google Sheets → Betha Cloud (lançamento)
```

---

## Funcionalidades

| # | Funcionalidade | Descrição |
|---|---------------|-----------|
| 1 | **Sincronização de Bases** | Atualiza tabelas de referência (CID, CRM, tipos de afastamento etc.) da API Betha para o Sheets |
| 2 | **Download e Importação SOC** | Baixa relatório do SOC, valida o Excel gerado e importa os dados para a planilha Google |
| 3 | **Exportação por Período** | Gera relatório SOC para um intervalo de datas e salva o Excel localmente; suporta retomada de exportações interrompidas |
| 4 | **Lançamento no Betha** | Lê os registros marcados na planilha e lança cada atestado automaticamente via API Betha |
| 5 | **Gerenciamento de Credenciais** | Salva de forma segura (criptografada) os acessos do Betha, SOC e Proxy direto no `settings.json` |
| 6 | **Renovação de Token** | Renova o token de acesso Betha sem precisar reabrir o sistema |
| 7 | **Notificações WhatsApp** | Envia alertas de status via CallMeBot ao final das automações |

---

## Tecnologias

| Tecnologia | Uso no projeto |
|-----------|---------------|
| **Python 3.x** | Linguagem principal |
| **customtkinter** | Interface gráfica moderna |
| **requests** | Chamadas HTTP para APIs (Betha, SOC, CallMeBot) |
| **pandas** | Leitura, validação e manipulação de planilhas Excel |
| **Google Sheets API** | Leitura e escrita de dados na planilha de controle |
| **API Betha Cloud** | Envio de atestados e consulta de dados cadastrais |
| **API SOC** | Download automatizado de relatórios |
| **CallMeBot API** | Notificações via WhatsApp |
| **PyInstaller** | Empacotamento em executável `.exe` |

---

## Estrutura do Projeto

```
Automcao_atestados/
│
├── config/                        # Configurações globais
│   ├── __init__.py
│   ├── loaders.py                 # Carrega e descriptografa o settings.json
│   └── settings.json              # ⚠️  Credenciais e parâmetros (não versionado)
│
├── repositories/                  # Acesso a dados externos
│   └── update_data.py             # Sincronização Betha → Sheets
│
├── services/                      # Lógica de negócio e integrações
│   ├── auth_service.py            # Autenticação e criptografia de credenciais
│   ├── betha_service.py           # Integração completa com a API Betha
│   ├── callmebot_service.py       # Envio de notificações WhatsApp
│   ├── sheets_service.py          # Leitura/escrita no Google Sheets
│   ├── soc_service.py             # Download de relatórios SOC
│   ├── utils_service.py           # Utilitários gerais (log, formatação etc.)
│   └── validation_service.py      # Validação e sanitização dos dados Excel
│
├── estrutura.py                   # Dataclasses e estruturas de dados compartilhadas
├── config_global.py               # Instâncias globais (ex: objeto `sheets`)
├── gui.py                         # ▶  Ponto de entrada — Interface gráfica
├── main.py                        # Ponto de entrada alternativo (CLI/legado)
├── .gitignore
└── README.md
```

> **Dica:** Para entender o fluxo da interface, consulte o mapa de botões no topo do `gui.py`.

---

## Pré-requisitos

- **Python 3.10+** instalado
- **Git** (para clonar o repositório)
- Acesso às APIs: Betha Cloud, SOC e Google Sheets (credenciais necessárias)
- Conta CallMeBot configurada (opcional, para notificações)

---

## Instalação

### 1. Clone o repositório

```bash
git clone https://github.com/caledcresst-dev/Automcao_atestados.git
cd Automcao_atestados
```

### 2. Crie e ative um ambiente virtual

```bash
python -m venv venv

# Windows
.\venv\Scripts\activate

# Linux / macOS
source venv/bin/activate
```

### 3. Instale as dependências

```bash
pip install customtkinter requests pandas \
            google-auth-oauthlib google-api-python-client \
            openpyxl pyinstaller
```

---

## Configuração

O projeto utiliza o arquivo `config/settings.json` para armazenar todos os parâmetros de acesso. **Esse arquivo não é versionado** — crie-o manualmente antes de executar o sistema.

### Criando o `settings.json`

```json
{
  "betha": {
    "api": {
      "base_url": "https://api.betha.cloud/...",
      "url_login": "https://login.betha.cloud/...",
      "endpoints": {
        "atestado":           "/atestados",
        "cid":                "/cids",
        "medico":             "/medicos",
        "tipo_afastamento":   "/tipos-afastamento",
        "tipo_atestado":      "/tipos-atestado",
        "motivo_consulta":    "/motivos-consulta",
        "pessoa_juridica":    "/pessoas-juridicas",
        "listagem_matricula": "/matriculas",
        "anexo":              "/anexos"
      },
      "user_access": "SEU_USER_ACCESS_TOKEN",
      "authorization": "Bearer SEU_TOKEN"
    },
    "user": {
      "admin": {
        "LOGIN":    "LOGIN_CRIPTOGRAFADO",
        "PASSWORD": "SENHA_CRIPTOGRAFADA"
      }
    }
  },
  "soc": {
    "URL_SOC": "https://sistema.soc.com.br/...",
    "user": {
      "admin": {
        "LOGIN":         "seu_email@empresa.com",
        "PASSWORD":      "sua_senha",
        "SENHA_VIRTUAL": "1,2,3,4"
      }
    }
  },
  "proxy": {
    "PROXY_HOST": "proxy.empresa.com",
    "PROXY_PORT": "8080",
    "PROXY_USER": "usuario_proxy",
    "PROXY_PASS":  "senha_proxy"
  },
  "google_sheets": {
    "planilha": "Nome da Planilha",
    "aba":      "Nome da Aba"
  },
  "paths": {
    "downloads": "C:/Users/usuario/Downloads/atestados",
    "logs":      "C:/Users/usuario/Downloads/atestados/logs"
  }
}
```

> **⚠️ Segurança:** Os campos de `LOGIN` e `PASSWORD` armazenam valores criptografados pelo `auth_service`. Ao salvar pelas **Credenciais de Acesso** na GUI, a criptografia é aplicada automaticamente — não preencha esses campos manualmente com senhas em texto puro.

### Autenticação Google Sheets

Para que o Sheets funcione, coloque o arquivo de credenciais da Service Account do Google (`credentials.json`) na raiz do projeto e compartilhe a planilha com o e-mail da service account.

---

## Como Usar

### Iniciando a interface

```bash
python gui.py
```

### Fluxo recomendado de uso

```
1. Na primeira execução → configure as Credenciais de Acesso (⚙️)
2. Clique em "Atualizar Token Betha" para garantir autenticação válida
3. Use "1. Atualizar Bases" para sincronizar CID, CRM e demais tabelas
4. Use "2. Baixar relatório SOC" para obter e importar os dados do período
5. Revise a planilha Google Sheets e marque os registros como "Sim"
6. Use "4. Lançar atestado no Betha" para enviar os atestados revisados
```

---

## Interface Gráfica

![Interface gráfica da aplicação](<Captura de tela 2026-02-12 120847.png>)

A interface é dividida em duas áreas principais:

**Barra lateral (esquerda) — Ações:**

| Botão | Ação |
|-------|------|
| 1. Atualizar Bases | Sincroniza CID, CRM e outros dados de consulta |
| 2. Baixar relatório SOC | Abre seletor de período e inicia o download + importação |
| 3. Exportar por período | Gera relatório SOC detalhado com opção de retomar exportações pausadas |
| 4. Lançar atestado no Betha | Lê a planilha e envia os atestados marcados |
| Limpar Console | Apaga o histórico de execução visível |
| Atualizar Token Betha | Renova o token de autenticação |
| ⚙️ Credenciais | Abre o painel de configuração de acessos |

**Painel central (direita) — Console:**

Exibe em tempo real o log de todas as operações com timestamps, emojis de status e mensagens de erro detalhadas.

---

## Geração de Executável

Para distribuir o sistema sem necessidade de Python instalado:

```bash
pyinstaller --noconfirm --onefile --windowed --add-data "config;config" --name "AutomacaoAtestados" gui.py
```

O executável será gerado em `dist/AutomacaoAtestados.exe`.

> **Atenção:** O arquivo `config/settings.json` **não é embutido** no executável por segurança. Distribua-o separadamente e oriente o usuário a colocá-lo na mesma pasta do `.exe`.

---

## Contribuição

1. Faça um fork do repositório
2. Crie uma branch para sua feature:
   ```bash
   git checkout -b feature/minha-feature
   ```
3. Faça commit das alterações:
   ```bash
   git commit -m "feat: descrição clara da mudança"
   ```
4. Envie para o repositório remoto:
   ```bash
   git push origin feature/minha-feature
   ```
5. Abra um **Pull Request** descrevendo o que foi alterado e por quê

---

## Licença

Este projeto está licenciado sob a **Licença MIT**. Consulte o arquivo `LICENSE` para mais detalhes.

---

## Autor

Desenvolvido por **Catrique** ([@catrique](https://github.com/catrique))