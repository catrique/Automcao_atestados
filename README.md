# 🏥 Automação de Atestados — CRESST

> Sistema em Python para automação completa do ciclo de gestão de atestados médicos, integrando **SOC**, **Google Sheets**, **Betha Cloud**.

---

## 📋 Índice

* [Sobre o Projeto](#sobre-o-projeto)
* [Fluxo do Sistema](#fluxo-do-sistema)
* [Funcionalidades Principais](#funcionalidades-principais)
* [Arquitetura e Componentes](#arquitetura-e-componentes)
* [Tecnologias Utilizadas](#tecnologias-utilizadas)
* [Estrutura do Projeto](#estrutura-do-projeto)
* [Pré-requisitos](#pré-requisitos)
* [Instalação](#instalação)
* [Configuração](#configuração)
* [Como Usar](#como-usar)
* [Interface Gráfica](#interface-gráfica)
* [Segurança e Criptografia](#segurança-e-criptografia)
* [Geração do Executável (.EXE)](#geração-do-executável-exe)
* [Licença](#licença)
* [Autor](#autor)

---

# 💡 Sobre o Projeto

O **Automação de Atestados — CRESST** foi desenvolvido para eliminar processos manuais e repetitivos na gestão administrativa de afastamentos médicos em ambientes públicos e corporativos.

A aplicação centraliza todo o fluxo operacional em uma única interface gráfica desenvolvida com `customtkinter`, permitindo:

* Extração automatizada de relatórios do sistema **SOC**
* Validação e sanitização de dados com **Pandas**
* Consolidação e auditoria via **Google Sheets**
* Envio automatizado de atestados para o **Betha Cloud**

O sistema foi projetado para operar em ambientes corporativos restritivos, incluindo intranets governamentais com proxy, firewall e limitações de certificados SSL.

---

# 🔄 Fluxo do Sistema

```text
[ Painel SOC ]
       │
       ▼
[ Selenium WebDriver ]
       │
       ▼
[ Validação e Higienização Pandas ]
       │
       ▼
[ Google Sheets API ]
       │
       ▼
[ Betha Cloud API ]
       │
       ▼
[ Auditoria ]
```

Fluxo operacional recomendado:

```text
1. Configurar credenciais
2. Atualizar token Betha
3. Sincronizar bases cadastrais
4. Baixar relatório SOC
5. Validar e revisar planilha
6. Enviar atestados ao Betha
7. Auditar resultados
```

---

# ⚙️ Funcionalidades Principais

| #  | Funcionalidade                          | Descrição                                                                                  |
| -- | --------------------------------------- | ------------------------------------------------------------------------------------------ |
| 1  | **Sincronização de Bases**              | Atualiza CID, CRM, vínculos, matrículas e demais tabelas da API Betha para o Google Sheets |
| 2  | **Download Automatizado SOC**           | Login automático, navegação e exportação de relatórios por período usando Selenium         |
| 3  | **Importação e Validação Excel**        | Processa planilhas via Pandas aplicando filtros regulatórios e sanitização                 |
| 4  | **Exportação por Período**              | Geração de relatórios detalhados com suporte à retomada de exportações interrompidas       |
| 5  | **Lançamento Automático Betha**         | Leitura da planilha Google e envio automático dos atestados via API REST                   |
| 6  | **Gerenciamento Seguro de Credenciais** | Armazenamento criptografado de acessos Betha, SOC e Proxy                                  |
| 7  | **Renovação de Token**                  | Atualização automática do token Betha sem reiniciar o sistema                              |
| 8  | **Auditoria Operacional**               | Registro de operador, IP e timestamp após integrações concluídas                           |                                       |
| 9 | **Compatibilidade Corporativa**         | Suporte a proxy corporativo, SSL customizado e ambientes restritos                         |

---

# 🏗️ Arquitetura e Componentes

A aplicação segue um modelo modular síncrono com utilização de **threads assíncronas** (`threading.Thread`) para impedir o congelamento da interface gráfica durante operações intensivas de I/O.

## Componentes Estruturais

### `config_global.py`

Centraliza serviços globais e mantém a conexão singleton com o Google Sheets.

### `config_network.py`

Aplica ajustes avançados de rede para ambientes com proxy corporativo e restrições SSL.

### `update_data.py`

Sincronizador responsável pela paginação da API Betha e política automática de retry.

### `auth_service.py`

Responsável pela criptografia local de credenciais e autenticação.

### `betha_service.py`

Gerencia payloads, autenticação e chamadas REST para a API Betha.

### `soc_service.py`

Executa automação web do sistema SOC utilizando Selenium.

### `validation_service.py`

Implementa validações regulatórias para CPF, CRM, CID e consistência dos dados.

### `sheets_service.py`

Wrapper de integração entre Pandas e Google Sheets.

### `utils_service.py`

Centraliza logs, tratamento de erros HTTP e integração com GUI.

---

# 🧰 Tecnologias Utilizadas

| Tecnologia                | Finalidade                    |
| ------------------------- | ----------------------------- |
| **Python 3.10+**          | Linguagem principal           |
| **customtkinter**         | Interface gráfica             |
| **requests**              | Comunicação HTTP              |
| **pandas**                | Manipulação de planilhas      |
| **Selenium WebDriver**    | Automação do SOC              |
| **Google Sheets API**     | Controle operacional          |
| **Betha Cloud API**       | Gestão de atestados           |
| **SOC API / Web**         | Extração de relatórios        |
| **PyInstaller**           | Empacotamento `.exe`          |
| **cryptography / Fernet** | Criptografia local            |
| **openpyxl**              | Manipulação de arquivos Excel |

---

# 📂 Estrutura do Projeto

```text
📂 automacao_atestados/
│
├── 📂 config/
│   ├── 📄 credentials.json       # Credencial Google Cloud Service Account
│   ├── 📄 loaders.py             # Gerenciador de configurações
│   └── 📄 settings.json          # Configurações e credenciais criptografadas
│
├── 📂 repositories/
│   └── 📄 update_data.py         # Sincronização Betha → Sheets
│
├── 📂 services/
│   ├── 📄 auth_service.py        # Criptografia e autenticação
│   ├── 📄 betha_service.py       # Integração API Betha
│   ├── 📄 sheets_service.py      # Integração Google Sheets
│   ├── 📄 soc_service.py         # Automação SOC
│   ├── 📄 utils_service.py       # Logs e utilitários
│   └── 📄 validation_service.py  # Validação de dados
│
├── 📄 config_global.py           # Serviços globais
├── 📄 config_network.py          # Configurações de rede
├── 📄 gui.py                     # Interface principal
├── 📄 README.md
└── 📄 .gitignore
```

---

# 📋 Pré-requisitos

## Ambiente Operacional

* Windows 10 ou superior
* Python 3.10+
* Google Chrome instalado
* Git instalado (opcional)

## Acessos Necessários

* API Betha Cloud
* Sistema SOC
* Google Sheets API
* Proxy corporativo (se aplicável)

---

# 🚀 Instalação

## 1. Acessar o repositório

```bash
cd Automcao_atestados
```
## 2. Instalar dependências

### Via requirements.txt

```bash
pip install -r requirements.txt
```

### Ou instalação manual

```bash
pip install customtkinter requests pandas \
            google-auth-oauthlib google-api-python-client \
            openpyxl selenium pyinstaller cryptography
```

---

# ⚙️ Configuração

O sistema utiliza o arquivo `config/settings.json` para armazenar parâmetros estruturais e credenciais criptografadas.

> ⚠️ O arquivo `settings.json` NÃO deve ser versionado.

## Exemplo de configuração

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
    "downloads": "pasta onde será baixado os arquivos"
  }
}
```

## Configuração Google Sheets

1. Criar uma Service Account no Google Cloud
2. Baixar o `credentials.json`
3. Colocar o arquivo dentro da pasta `config/`
4. Compartilhar a planilha com o e-mail da Service Account

---

# ▶️ Como Usar

## Inicialização

```bash
python gui.py
```

## Fluxo recomendado

```text
1. Configurar Credenciais (⚙️)
2. Atualizar Token Betha
3. Atualizar Bases
4. Baixar Relatório SOC
5. Revisar Google Sheets
6. Lançar Atestados
```

---

# 🖥️ Interface Gráfica

A interface é dividida em duas regiões principais:

## Barra Lateral — Ações

| Botão                | Função                             |
| -------------------- | ---------------------------------- |
| Atualizar Bases      | Sincroniza tabelas cadastrais      |
| Baixar Relatório SOC | Download e importação do relatório |
| Exportar por Período | Exporta relatórios detalhados      |
| Lançar no Betha      | Envia registros aprovados          |
| Atualizar Token      | Renova autenticação Betha          |
| Limpar Console       | Limpa logs visíveis                |
| ⚙️ Credenciais       | Configuração de acessos            |

## Painel Central — Console

Exibe logs operacionais em tempo real:

* Status das integrações
* Progresso das tarefas
* Erros HTTP
* Eventos de auditoria
* Logs de importação/exportação

---

# 🔐 Segurança e Criptografia

O sistema implementa proteção local baseada em hardware utilizando criptografia simétrica (`Fernet/AES`).

## Recursos de segurança

* Criptografia de credenciais
* Associação ao identificador físico da máquina
* Proteção contra cópia simples de arquivos
* Armazenamento seguro de tokens
* Separação de credenciais operacionais

> ⚠️ Nunca versionar `settings.json` ou `credentials.json`.

---

# 📦 Geração do Executável (.EXE)

Para distribuição sem necessidade de Python instalado:

```bash
pyinstaller --noconfirm --onefile --windowed \
--add-data "config;config" \
--name "AutomacaoAtestados" gui.py
```

O executável será gerado em:

```text
dist/AutomacaoAtestados.exe
```

## Observações importantes

* `settings.json` NÃO é embutido no executável
* Distribuir credenciais separadamente
* Chrome deve estar instalado na máquina cliente

---

# 🛠️ Troubleshooting e Limitações

| Problema            | Possível causa                                |
| ------------------- | --------------------------------------------- |
| Falha no login SOC  | Credenciais inválidas ou bloqueio corporativo |
| Timeout Betha       | Instabilidade de rede ou token expirado       |
| Erro Google Sheets  | Service Account sem permissão                 |
| Interface travando  | Operação fora de thread assíncrona            |
| Selenium não inicia | Chrome incompatível ou bloqueado              |

## Recomendações

* Validar proxy antes da execução
* Garantir acesso externo às APIs
* Atualizar Google Chrome regularmente
* Revisar permissões da planilha Google
* Executar logs em modo administrativo quando necessário

---

# 👨‍💻 Autor

Desenvolvido por **Cáled Tarique**.
