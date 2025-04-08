# DUA Automation

<p align="center">
  <img src="resources/logo_new.png" alt="DUA Automation Logo" width="200"/>
</p>

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green)](LICENSE)

**DUA Automation** é uma ferramenta para automatizar a emissão de Documentos Únicos de Arrecadação (DUAs) para o estado do Espírito Santo, eliminando a necessidade de preenchimento manual do formulário na página da SEFAZ-ES.

## ✨ Funcionalidades

- **Processamento em lote**: Emita múltiplos DUAs a partir de um arquivo CSV ou Excel
- **Resolução automática de CAPTCHA**: Utiliza técnicas avançadas para resolver CAPTCHAs automaticamente
- **Interface gráfica amigável**: Fácil de usar, mesmo para usuários não técnicos
- **Salva PDFs automaticamente**: Todos os DUAs gerados são salvos organizadamente em formato PDF
- **Compatibilidade multiplataforma**: Funciona em Windows, macOS e Linux

## 📋 Pré-requisitos

- Python 3.8 ou superior
- Google Chrome instalado (versão 90 ou superior)
- FFmpeg instalado (obrigatório para processamento de áudio do CAPTCHA)
- Conexão com a internet

### Instalação do FFmpeg

#### Windows
- Baixe o FFmpeg do [site oficial](https://ffmpeg.org/download.html) ou use o [Chocolatey](https://chocolatey.org/):
  ```
  choco install ffmpeg
  ```
- Ou siga este tutorial: [Como instalar FFmpeg no Windows](https://www.wikihow.com/Install-FFmpeg-on-Windows)

#### macOS
- Usando Homebrew:
  ```
  brew install ffmpeg
  ```

#### Linux
- Ubuntu/Debian:
  ```
  sudo apt update
  sudo apt install ffmpeg
  ```
- Fedora:
  ```
  sudo dnf install ffmpeg
  ```
- Arch Linux:
  ```
  sudo pacman -S ffmpeg
  ```

## 🔧 Instalação

### Usando o instalador (Windows)

1. Baixe o arquivo `DUA_Automation_Setup.exe` da [página de releases](https://github.com/bielaltoe/dua_automation/releases)
2. Execute o instalador e siga as instruções
3. Após a instalação, inicie o programa pelo atalho no Menu Iniciar

### Instalação manual (todas as plataformas)

```bash
# Clonar o repositório
git clone https://github.com/bielaltoe/dua_automation.git
cd dua_automation

# Criar ambiente virtual
python -m venv venv

# Ativar ambiente virtual
# No Windows:
venv\Scripts\activate
# No macOS/Linux:
source venv/bin/activate

# Instalar dependências
pip install -r requirements.txt

# Executar o programa
python run_ui.py
```

## 📊 Como usar

1. **Prepare seu arquivo de dados**:
   - Crie um arquivo CSV ou Excel com as seguintes colunas:
     - `CPF/CNPJ`: número do contribuinte
     - `SERVIÇO`: código do serviço (ex: 138-4)
     - `REFERENCIA`: mês/ano de referência (ex: 01/2024)
     - `VENCIMENTO`: data de vencimento (ex: 10/01/2024)
     - `VALOR`: valor do DUA (ex: 123,45)
     - `NOTA FISCAL`: número da nota fiscal (opcional)
     - `INFORMAÇÕES ADICIONAIS`: texto adicional (opcional)
   - Baixe um modelo de arquivo: [CSV](https://github.com/bielaltoe/dua_automation/raw/data.csv)/[XLSX](https://github.com/bielaltoe/dua_automation/raw/dua_excel.xlsx)

2. **Execute o aplicativo DUA Automation**

3. **Carregue seu arquivo de dados**:
   - Clique em "Selecionar Arquivo"
   - Escolha o arquivo CSV ou Excel que você preparou
   - Verifique se os dados foram carregados corretamente na tabela

4. **Configure o diretório de saída para os PDFs**:
   - Navegue até a aba "Configurações"
   - Defina o diretório onde os PDFs serão salvos

5. **Inicie o processamento**:
   - Clique em "Iniciar Processamento"
   - O sistema abrirá o Chrome automaticamente e começará a processar cada registro
   - Os PDFs gerados serão salvos no diretório escolhido

6. **Resolução de CAPTCHAs**:
   - O sistema tentará resolver os CAPTCHAs automaticamente
   - Se necessário, auxiliará você a resolver CAPTCHAs manualmente

## 🚀 Exemplos

### Formato do arquivo CSV
```csv
CPF/CNPJ,SERVIÇO,REFERENCIA,VENCIMENTO,VALOR,NOTA FISCAL,INFORMAÇÕES ADICIONAIS
27277961000102,138-4,02/2025,15/02/2025,429.80,14539657,Rio Branco Alimentos S/A
27277961000102,138-4,02/2025,15/02/2025,1.00,32132123,teste 1
```

### Códigos de serviço suportados
- `138-4`: ICMS - Substituição Tributaria - Contribuintes sediados no ES
- `137-6`: ICMS - Substituição Tributária - Contribuintes sediados fora do ES
- `386-7`: ICMS - Diferencial de Alíquota EC 87
- `121-0`: ICMS - Comércio
- E outros (consulte o código para lista completa)

## 🔄 Processo de build

Para compilar o DUA Automation em um executável:

### Windows
```batch
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
python build.py
```

### Linux/macOS
```bash
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
python build.py
```

## 📃 Licença

Este projeto está licenciado sob a [Licença MIT](LICENSE) - veja o arquivo LICENSE para detalhes.

## 🤝 Contribuindo

Contribuições são bem-vindas! Por favor, leia o arquivo [CONTRIBUTING.md](CONTRIBUTING.md) para detalhes sobre nosso código de conduta e o processo para enviar pull requests.

## 🙏 Agradecimentos

- [RecaptchaBypass](https://github.com/obaskly/RecaptchaBypass) por fornecer a base para a solução de CAPTCHA
- Todos os contribuidores que ajudaram no desenvolvimento

