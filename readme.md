
# 📊 Extração de Leads Diária

Aplicação em Python para extração de leads a partir de um banco de dados SQL Server, com base em uma data específica, exportando os resultados para um arquivo Excel (.xlsx).

O aplicativo funciona via menu interativo no terminal e utiliza um arquivo de configuração externo (.ini) para armazenar credenciais e a query SQL.


## 🚀 Funcionalidades

- 📅 Extrair leads do dia anterior
- 🗓️ Extrair leads de uma data específica
- 📄 Exportação automática para Excel (.xlsx)
- 🔐 Configurações e query SQL separadas do código
- 📦 Possibilidade de gerar executável .exe (sem necessidade de Python instalado)


 ## 🧱 Tecnologias Utilizadas

- Python 3.9+
- pyodbc
- pandas
- openpyxl
- configparser
- SQL Server
- ODBC Driver 17 for SQL Server
- PyInstaller (para gerar o .exe)


## 📁 Estrutura do Projeto
📦 projeto  
 ┣ 📂 config  
 ┃ ┗ 📄 config.ini  
 ┣ 📄 main.py  
 ┣ 📄 README.md  


## ⚙️ Configuração do Ambiente (Desenvolvimento)
### 1️⃣ Criar ambiente virtual (opcional, recomendado)
python -m venv venv
venv\Scripts\activate

### 2️⃣ Instalar dependências
pip install pyodbc pandas openpyxl pyinstaller

## 🛠️ Arquivo de Configuração (config.ini)

Crie o arquivo config/config.ini com o seguinte conteúdo:

[database]  
host=SEU_SERVIDOR  
port=1433  
database=SEU_BANCO  
user=SEU_USUARIO  
password=SUA_SENHA  

[query]
find_leads=SELECT * FROM leads WHERE CAST(data_criacao AS DATE) = ?


**📌 Observações importantes:** 

- A query deve conter ? para receber a data
- O formato da data utilizado é YYYY-MM-DD

## ▶️ Executar em Modo Python

No diretório do projeto:

python main.py


## 📤 Saída Gerada

- Arquivo Excel criado na pasta onde o app é executado
- Nome do arquivo:
leads_YYYY-MM-DD.xlsx

#### Exemplo: leads_2024-01-15.xlsx

## 🧩 Gerando o Executável (.exe) com PyInstaller
✅ Pré-requisitos

- Windows
- Python instalado
- ODBC Driver 17 for SQL Server instalado na máquina de destino
- Dependências já instaladas no ambiente


---
### 1️⃣ Instalar o PyInstaller
pip install pyinstaller

### 2️⃣ Comando para Gerar o .exe

Execute na raiz do projeto:

pyinstaller --onefile --name ExtracaoLeads main.py


### 📌 Isso irá:

Gerar um único arquivo .exe

Criar as pastas build/ e dist/

### 3️⃣ Estrutura Após Build  
📦 projeto  
 ┣ 📂 build  
 ┣ 📂 dist  
 ┃ ┗ 📄 ExtracaoLeads.exe  
 ┣ 📂 config  
 ┃ ┗ 📄 config.ini  
 ┣ 📄 main.py  
 ┣ 📄 README.md  

--- 
**⚠️ O arquivo config.ini deve estar junto da pasta config, no mesmo diretório do .exe.**

### Exemplo de distribuição final:

📦 ExtracaoLeads  
 ┣ 📂 config  
 ┃ ┗ 📄 config.ini  
 ┗ 📄 ExtracaoLeads.exe  

### 4️⃣ Executar o .exe

Basta dar duplo clique em:  
ExtracaoLeads.exe

Ou executar via terminal:  
ExtracaoLeads.exe

## ⚠️ Problemas Comuns
### ❌ Erro de conexão com SQL Server

- Verifique:
  - Host, usuário e senha no config.ini
  - Porta correta
  - Driver ODBC instalado

### ❌ Tela preta abrindo e fechando

Execute o .exe via Prompt de Comando

Assim será possível ver mensagens de erro

## 🔒 Segurança

- Nenhuma credencial fica hardcoded no código
- Informações sensíveis ficam no config.ini
- Uso de parâmetros na query evita SQL Injection



## 👨‍💻 Autor

Desenvolvido por Pietro Stortti Venzon
📌 Projeto para automação de extração de leads