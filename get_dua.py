from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import csv
import time
import os
import pandas as pd
import requests
from urllib.parse import urlparse

# Import the RecaptchaSolver
from RecaptchaBypass.RecaptchaSolver import RecaptchaSolver
from pathlib import Path

# Add webdriver_manager to automatically download the correct ChromeDriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import sys
import platform
import subprocess
import traceback
import tempfile
import zipfile
import urllib.request

# Fix ChromeType import issue - handle different webdriver-manager versions
try:
    # For newer versions of webdriver-manager
    from webdriver_manager.chrome import ChromeDriverManager, ChromeType
except ImportError:
    # For older versions where ChromeType might be defined differently
    from webdriver_manager.chrome import ChromeDriverManager

    # Create a fallback ChromeType enum
    class ChromeType:
        GOOGLE = "google"
        CHROMIUM = "chromium"


# Configurações
CSV_PATH = "dados.csv"

# Constants for Chrome portable
CHROME_PORTABLE_VERSION = "114.0.5735.90"  # A stable Chrome version
CHROME_PORTABLE_URL = {
    "Windows": f"https://storage.googleapis.com/chrome-for-testing-public/{CHROME_PORTABLE_VERSION}/win64/chrome-win64.zip",
    "Darwin": f"https://storage.googleapis.com/chrome-for-testing-public/{CHROME_PORTABLE_VERSION}/mac-x64/chrome-mac-x64.zip",  # For macOS
    "Linux": f"https://storage.googleapis.com/chrome-for-testing-public/{CHROME_PORTABLE_VERSION}/linux64/chrome-linux64.zip",
}


# Determine appropriate PDF directory
def get_pdf_directory():
    """Get an appropriate PDF directory that works for both development and executable environments"""
    # First try to get from environment variable (set by run_ui.py)
    if "PDF_DIR" in os.environ:
        pdf_dir = os.environ["PDF_DIR"]
    else:
        # For executable, use a folder next to the executable or in user's documents
        if getattr(sys, "frozen", False):
            # Running as executable
            try:
                # Try to use a directory next to the executable
                exe_dir = os.path.dirname(sys.executable)
                pdf_dir = os.path.join(exe_dir, "pdfs_gerados")
            except:
                # Fallback to user's documents folder
                import pathlib

                user_docs = os.path.join(pathlib.Path.home(), "Documents")
                pdf_dir = os.path.join(user_docs, "DUA_Automation", "pdfs_gerados")
        else:
            # Running in development
            pdf_dir = os.path.abspath("pdfs_gerados")

    # Ensure directory exists
    os.makedirs(pdf_dir, exist_ok=True)
    return pdf_dir


# Set PDF directory
PDF_DIR = get_pdf_directory()
print(f"PDFs will be saved to: {PDF_DIR}")

# Configurar opções do Chrome
options = webdriver.ChromeOptions()
options.add_argument("--disable-dev-shm-usage")
options.add_argument("--no-sandbox")
options.add_argument("--log-level=3")
options.add_argument("--no-proxy-server")
options.add_argument("--incognito")
# Forçar idioma inglês para compatibilidade com o solver de CAPTCHA
options.add_argument("--lang=en-US")
options.add_argument("--language=en-US")
options.add_experimental_option(
    "prefs",
    {
        "download.default_directory": PDF_DIR,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,  # Isso faz com que PDFs sejam baixados em vez de abertos
        "download.open_pdf_in_system_reader": False,
        # Adicionar configurações de idioma aqui também
        "intl.accept_languages": "en-US,en",
        "profile.default_content_setting_values.geolocation": 2,
    },
)
options.add_experimental_option(
    "excludeSwitches", ["enable-automation", "enable-logging"]
)
options.add_experimental_option("useAutomationExtension", False)

# Adicionar cabeçalho Accept-Language para requisições
options.add_argument("--accept-lang=en-US,en;q=0.9")

# Variável global para o driver - inicialmente None
driver = None

# Variável global para controle de interrupção
stop_requested = False

# Adicionar variáveis globais para comunicação com a UI
captcha_callback = None
manual_captcha_requested = False


def set_stop_flag():
    """Set a global stop flag to interrupt any ongoing operations"""
    global stop_requested
    stop_requested = True
    print("Stop flag set in get_dua module")


def check_stop_flag():
    """Check if stop was requested"""
    return stop_requested


def reset_stop_flag():
    """Reset the stop flag"""
    global stop_requested
    stop_requested = False


def set_captcha_callback(callback_function):
    """Define a callback function to be called when manual CAPTCHA solving is needed"""
    global captcha_callback
    captcha_callback = callback_function


def get_portable_chrome_path():
    """Get the path to the portable Chrome executable"""
    system = platform.system()
    base_dir = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "chrome-portable"
    )

    if system == "Windows":
        chrome_path = os.path.join(base_dir, "chrome-win64", "chrome.exe")
    elif system == "Darwin":  # macOS
        chrome_path = os.path.join(
            base_dir,
            "chrome-mac-x64",
            "Google Chrome for Testing.app",
            "Contents",
            "MacOS",
            "Google Chrome for Testing",
        )
    else:  # Linux
        chrome_path = os.path.join(base_dir, "chrome-linux64", "chrome")

    return chrome_path


def download_portable_chrome():
    """Download and extract portable Chrome if not already available"""
    system = platform.system()
    if system not in CHROME_PORTABLE_URL:
        print(f"Unsupported system for portable Chrome: {system}")
        return None

    base_dir = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "chrome-portable"
    )
    chrome_path = get_portable_chrome_path()

    # Check if Chrome already exists
    if os.path.exists(chrome_path):
        print(f"Portable Chrome already exists at: {chrome_path}")
        return chrome_path

    # Create base directory
    os.makedirs(base_dir, exist_ok=True)

    # Download Chrome
    try:
        print(f"Downloading portable Chrome {CHROME_PORTABLE_VERSION}...")
        chrome_url = CHROME_PORTABLE_URL[system]

        # Create a temporary file for the download
        with tempfile.NamedTemporaryFile(delete=False, suffix=".zip") as temp_file:
            zip_path = temp_file.name

        # Download the Chrome zip file
        urllib.request.urlretrieve(chrome_url, zip_path)

        # Extract the zip file
        print(f"Extracting portable Chrome to {base_dir}...")
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(base_dir)

        # Make the Chrome binary executable on Unix systems
        if system != "Windows":
            os.chmod(chrome_path, 0o755)

        # Clean up the temporary file
        os.unlink(zip_path)

        print(f"Portable Chrome installed at: {chrome_path}")
        return chrome_path

    except Exception as e:
        print(f"Error downloading portable Chrome: {e}")
        traceback.print_exc()
        return None


# Função para inicializar o WebDriver quando necessário
def initialize_driver():
    global driver
    if driver is None:
        print("Inicializando o navegador Chrome...")
        try:
            # Always try to get/use portable Chrome for stability
            portable_chrome_path = get_portable_chrome_path()

            # If portable Chrome doesn't exist, download it
            if not os.path.exists(portable_chrome_path):
                print("Portable Chrome not found, downloading...")
                portable_chrome_path = download_portable_chrome()

            if portable_chrome_path and os.path.exists(portable_chrome_path):
                print(f"Using portable Chrome from: {portable_chrome_path}")
                options.binary_location = portable_chrome_path
            else:
                print("Portable Chrome not available, falling back to system Chrome")
                chrome_path = find_chrome_executable()
                if chrome_path:
                    options.binary_location = chrome_path
                    print(f"Using system Chrome: {chrome_path}")

            # Tentar inicializar o Chrome com WebDriverManager automaticamente
            try:
                print("Tentando usar WebDriverManager para obter ChromeDriver...")
                # Use Chrome for Testing driver for better compatibility
                # Handle different webdriver-manager versions
                try:
                    service = Service(
                        ChromeDriverManager(chrome_type=ChromeType.CHROMIUM).install()
                    )
                except (TypeError, AttributeError):
                    # Fallback for older versions that don't support chrome_type
                    service = Service(ChromeDriverManager().install())

                driver = webdriver.Chrome(service=service, options=options)
                print("Chrome iniciado com WebDriverManager")
            except Exception as webdriver_error:
                print(f"Erro com WebDriverManager: {webdriver_error}")
                print("Tentando inicializar o Chrome diretamente...")
                driver = webdriver.Chrome(options=options)
                print("Chrome iniciado diretamente")

            # Abrir uma página padrão inicial
            driver.get("https://internet.sefaz.es.gov.br/agenciavirtual/")
            print("Navegador iniciado com sucesso.")

        except Exception as e:
            error_message = f"Erro ao inicializar Chrome: {str(e)}\n"
            error_message += traceback.format_exc()
            print(error_message)

            # Tentar com configurações alternativas
            try:
                print("Tentando configuração alternativa...")
                alt_options = webdriver.ChromeOptions()
                alt_options.add_argument("--headless=new")
                alt_options.add_argument("--disable-gpu")
                alt_options.add_argument("--no-sandbox")
                alt_options.add_argument("--disable-dev-shm-usage")

                service = Service(ChromeDriverManager().install())
                driver = webdriver.Chrome(service=service, options=alt_options)
                print("Chrome iniciado em modo alternativo")
                driver.get("https://internet.sefaz.es.gov.br/agenciavirtual/")
            except Exception as alt_error:
                print(f"Erro na configuração alternativa: {alt_error}")
                print(
                    "Falha ao inicializar o Chrome. Verifique se o Chrome está instalado corretamente."
                )
                # Propagar o erro para ser tratado pela UI
                raise Exception(f"Não foi possível inicializar o Chrome: {str(e)}")

    return driver


# Mapeamento dos códigos de serviço
# Formato: 'código no CSV': 'valor no dropdown HTML'
SERVICO_MAPPING = {
    "138-4": "1464",  # ICMS - Substituição Tributaria - Contribuintes sediados no ES
    "137-6": "1463",  # ICMS - Substituição Tributária - Contribuintes sediados fora do ES
    "386-7": "1439",  # ICMS - Diferencial de Alíquota EC 87
    "121-0": "1434",  # ICMS - Comércio
    "128-7": "1440",  # ICMS - Diferencial de Alíquota de Empresas Comerciais
    "129-5": "1443",  # ICMS - Diferencial de Alíquota de Empresas Industriais
    "125-2": "1462",  # ICMS - Serviços de Transporte - Empresas do Estado do Espírito Santo
    "122-8": "1455",  # ICMS - Indústria
    "145-7": "1437",  # ICMS - Demais Produtos
    "162-7": "1452",  # ICMS - Fundo Estadual de Combate a Pobreza
}


def preencher_formulario(dados):
    # Garantir que o driver está inicializado
    driver = initialize_driver()

    driver.get(
        "https://internet.sefaz.es.gov.br/agenciavirtual/area_publica/e-dua/icms.php"
    )

    # Check stop flag
    if stop_requested:
        print("Interrupção solicitada durante preenchimento do formulário")
        return False

    # Preencher campos
    driver.find_element(By.NAME, "codCpfCnpjPessoa").send_keys(dados["CPF_CNPJ"])

    # Verificar interrupção a cada passo importante
    if stop_requested:
        return False

    # Mapear o código do serviço para o valor do dropdown
    servico_codigo = dados["SERVICO"]
    servico_valor = SERVICO_MAPPING.get(servico_codigo, servico_codigo)

    # Usar Select para o campo de serviço (dropdown)
    servico_select = Select(driver.find_element(By.NAME, "idServico"))
    try:
        servico_select.select_by_value(servico_valor)
        print(f"Selecionado serviço: código {servico_codigo} -> valor {servico_valor}")
    except Exception as e:
        print(f"Erro ao selecionar serviço {servico_codigo}: {str(e)}")
        # Tentar selecionar pelo texto visível contendo o código
        for option in servico_select.options:
            if servico_codigo in option.text:
                option.click()
                print(f"Selecionado via texto: {option.text}")
                break

    if stop_requested:
        return False

    driver.find_element(By.NAME, "datReferencia").send_keys(dados["REFERENCIA"])
    driver.find_element(By.NAME, "datVencimento").send_keys(dados["VENCIMENTO"])
    driver.find_element(By.NAME, "vlrReceita").send_keys(dados["VALOR"])
    driver.find_element(By.NAME, "dscInformacao").send_keys(dados["INFO_COMBINADA"])

    # Melhorar a resolução do CAPTCHA com mais informações de diagnóstico
    # e compatibilidade entre Windows 10 e 11
    try:
        print("\nTentando resolver CAPTCHA automaticamente...")
        t0 = time.time()

        # Verificar a versão do Windows para ajustar comportamento
        win_version = platform.release()
        print(f"Sistema operacional: Windows {win_version}")

        # Certificar que a página está totalmente carregada antes de tentar resolver o CAPTCHA
        WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, "btnEnviar"))
        )

        # Aumentar o tempo de espera para elementos do CAPTCHA no Windows 10
        if win_version.startswith("10"):
            print(
                "Detectado Windows 10 - ajustando parâmetros do resolvedor de CAPTCHA"
            )
            # Scroll para garantir que o CAPTCHA esteja visível (problema comum no Windows 10)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
            time.sleep(1)  # Pequena pausa para garantir que o scroll foi concluído

        # Inicializar o resolvedor de CAPTCHA com mais opções de debug
        recaptchaSolver = RecaptchaSolver(driver, debug_mode=True)

        # Para Windows 10, tente a resolução com retry
        if win_version.startswith("10"):
            captcha_solved = False
            max_attempts = 3
            for attempt in range(1, max_attempts + 1):
                try:
                    print(f"Tentativa {attempt}/{max_attempts} de resolver CAPTCHA...")
                    recaptchaSolver.solveCaptcha()
                    captcha_solved = True
                    break
                except Exception as retry_error:
                    print(f"Falha na tentativa {attempt}: {str(retry_error)}")
                    # Pequena pausa entre tentativas
                    time.sleep(2)
        else:
            # Para Windows 11 e outros, comportamento normal
            recaptchaSolver.solveCaptcha()
            captcha_solved = True

        print(f"CAPTCHA resolvido em {time.time() - t0:.2f} segundos")
    except Exception as e:
        print(f"Falha na resolução automática do CAPTCHA: {str(e)}")
        print("Detalhes do erro:")
        traceback.print_exc()
        captcha_solved = False

        # Tentar tirar um screenshot do CAPTCHA para diagnóstico
        try:
            captcha_screenshot_path = os.path.join(
                PDF_DIR, f"captcha_error_{time.time()}.png"
            )
            driver.save_screenshot(captcha_screenshot_path)
            print(f"Screenshot do CAPTCHA salvo em: {captcha_screenshot_path}")
        except:
            pass

    # Se a resolução automática falhar, solicitar intervenção manual
    if not captcha_solved:
        global manual_captcha_requested
        manual_captcha_requested = True

        print("\n\nFalha na resolução automática do CAPTCHA!")
        print("Será necessário resolver o CAPTCHA manualmente.")

        if captcha_callback:
            # Chamar o callback para notificar a UI
            print("Notificando interface para intervenção manual...")
            captcha_callback()

            # Esperar até que o usuário sinalize que o CAPTCHA foi resolvido manualmente
            # ou até que o processo seja interrompido
            while manual_captcha_requested and not stop_requested:
                time.sleep(0.5)

            if stop_requested:
                return False
        else:
            # Se não há callback registrado (modo terminal), cai no modo antigo
            print("\n*********************************************")
            print("* RESOLVA O CAPTCHA MANUALMENTE NO NAVEGADOR *")
            print("*    Pressione ENTER após concluir          *")
            print("*********************************************")
            try:
                input()  # Aguarda intervenção manual
            except:
                # Se não houver terminal, espera um tempo fixo
                print(
                    "Nenhum terminal detectado, aguardando 30 segundos para resolução manual..."
                )
                time.sleep(30)

    if stop_requested:
        return False

    # Submeter formulário
    driver.find_element(By.ID, "btnEnviar").click()
    return True


def captcha_solved_signal():
    """Método para sinalizar que o CAPTCHA foi resolvido manualmente"""
    global manual_captcha_requested
    manual_captcha_requested = False
    print("Captcha foi resolvido manualmente pelo usuário")


def baixar_pdf(cpf_cnpj, referencia, observacao=None, valor=None):
    """
    Baixa o PDF do DUA.

    Args:
        cpf_cnpj: CPF ou CNPJ do contribuinte
        referencia: Período de referência (ex: '01/2024')
        observacao: Informações adicionais (opcional)
        valor: Valor do DUA (opcional)

    Returns:
        bool: True se o PDF foi baixado com sucesso, False caso contrário
    """
    # Garantir que o driver está inicializado
    driver = initialize_driver()

    try:
        # Aguardar até que o botão "Gerar DUA" esteja visível e clicável
        print("Aguardando botão 'Gerar DUA'...")

        # Modificar para usar menor tempo de espera e checar interrupção
        total_wait = 30  # segundos
        check_interval = 1  # verificar cada 1 segundo
        elapsed = 0

        while elapsed < total_wait:
            if stop_requested:
                print("Interrupção solicitada durante espera pelo botão Gerar DUA")
                return False

            try:
                gerar_dua_button = WebDriverWait(driver, check_interval).until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            "//button[contains(text(), 'Gerar DUA') or contains(@onclick, 'gerarDua')]",
                        )
                    )
                )
                break  # Botão encontrado, sair do loop
            except:
                elapsed += check_interval
                if elapsed >= total_wait:
                    raise TimeoutError(
                        "Tempo esgotado esperando pelo botão 'Gerar DUA'"
                    )

        print("Botão 'Gerar DUA' encontrado, clicando...")
        time.sleep(1)
        gerar_dua_button.click()

        if stop_requested:
            print("Interrupção solicitada após clicar no botão Gerar DUA")
            return False

        # Aguardar até que o botão "Imprimir ou Salvar PDF" esteja visível
        print("Aguardando botão 'Imprimir ou Salvar PDF'...")
        imprimir_button = WebDriverWait(driver, 30).until(
            EC.element_to_be_clickable(
                (
                    By.XPATH,
                    "//a[contains(text(), 'Imprimir ou Salvar PDF') or contains(@href, 'imprimir-dua.php')]",
                )
            )
        )

        if stop_requested:
            print(
                "Interrupção solicitada durante espera pelo botão Imprimir ou Salvar PDF"
            )
            return False

        # Obter o link da página HTML
        html_link = imprimir_button.get_attribute("href")
        print(f"Link da página encontrado: {html_link}")

        # Criar um nome para o arquivo PDF baseado nos dados
        # Formato: CPF-CNPJ_REF_VALOR_OBS.pdf
        # Limpar caracteres inválidos do observacao
        obs_parte = ""
        if observacao:
            # Limitar tamanho e remover caracteres inválidos para nome de arquivo
            obs_limpo = "".join(c for c in observacao if c.isalnum() or c in " -_")
            obs_limpo = obs_limpo.replace(" ", "_")[
                :30
            ]  # Limitar tamanho e substituir espaços
            if obs_limpo:
                obs_parte = f"_{obs_limpo}"

        # Adicionar valor se disponível
        valor_parte = ""
        if valor:
            valor_str = str(valor).replace(".", ",")
            valor_parte = f"_{valor_str}"

        # Criar nome do arquivo com todas as informações disponíveis
        pdf_filename = (
            f"{cpf_cnpj}_{referencia.replace('/', '_')}{valor_parte}{obs_parte}.pdf"
        )
        pdf_path = os.path.join(PDF_DIR, pdf_filename)
        path = Path(pdf_path)

        # Abrir a página HTML em uma nova aba
        driver.execute_script(f"window.open('{html_link}', '_blank');")

        # Mudar para a nova aba
        driver.switch_to.window(driver.window_handles[-1])

        # Aguardar o carregamento da página
        print("Aguardando carregamento da página HTML...")
        WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, "body"))
        )

        if stop_requested:
            print("Interrupção solicitada durante carregamento da página HTML")
            return False

        # Usar o CDP (Chrome DevTools Protocol) para gerar o PDF
        print("Convertendo página HTML em PDF...")
        pdf_params = {
            "printBackground": True,
            "preferCSSPageSize": True,
            "marginTop": 0,
            "marginBottom": 0,
            "marginLeft": 0,
            "marginRight": 0,
        }

        # Executar o comando de impressão via CDP
        pdf_data = driver.execute_cdp_cmd("Page.printToPDF", pdf_params)

        # Decodificar os dados PDF de base64
        import base64

        pdf_bytes = base64.b64decode(pdf_data["data"])

        # Salvar o PDF
        path.write_bytes(pdf_bytes)
        print(f"PDF gerado e salvo com sucesso: {pdf_path}")

        # Fechar a aba atual e voltar para a anterior
        driver.close()
        driver.switch_to.window(driver.window_handles[0])

        return True

    except Exception as e:
        print(f"Erro ao gerar PDF: {str(e)}")
        # Salvar screenshot quando ocorrer erro
        try:
            # Usar informações completas no nome do screenshot de erro também
            screenshot_name = f"erro_{cpf_cnpj}_{referencia.replace('/', '_')}{valor_parte}{obs_parte}.png"
            driver.save_screenshot(f"{PDF_DIR}/{screenshot_name}")
            print(f"Screenshot de erro salvo em {PDF_DIR}/{screenshot_name}")
        except:
            pass
        return False


def wait_for_download(directory, timeout=30):
    """Aguarda até que um download seja concluído no diretório especificado."""
    seconds = 0
    dl_wait = True

    while dl_wait and seconds < timeout:
        time.sleep(1)
        dl_wait = False
        files = os.listdir(directory)

        for fname in files:
            if fname.endswith(".crdownload"):
                dl_wait = True

        seconds += 1

    # Verifica se algum arquivo foi baixado nos últimos seconds
    return seconds < timeout


# Função para fechar o navegador
def close_browser():
    global driver
    if driver is not None:
        print("Fechando o navegador...")
        driver.quit()
        driver = None
        print("Navegador fechado com sucesso.")


# When directly running the script (not from UI)
if __name__ == "__main__":
    # Inicializar o driver apenas quando o script é executado diretamente
    initialize_driver()

    # Ler o CSV ou Excel com as configurações corretas
    try:
        # Verificar extensão do arquivo
        file_ext = os.path.splitext(CSV_PATH)[1].lower()

        if file_ext in [".xlsx", ".xls"]:
            # Arquivo Excel
            print(f"Detectado arquivo Excel: {CSV_PATH}")
            try:
                data = pd.read_excel(CSV_PATH, dtype=str)
                print("Arquivo Excel carregado com sucesso")
            except Exception as e:
                print(f"Erro ao carregar arquivo Excel: {str(e)}")
                sys.exit(1)
        else:
            # Arquivo CSV
            print(f"Detectado arquivo CSV: {CSV_PATH}")

            # Verificar se tem linha de comentário
            try:
                with open(CSV_PATH, "r", errors="ignore") as f:
                    first_line = f.readline().strip()
                    skip_rows = 1 if first_line.startswith("//") else 0

                # Tentar detectar o delimitador
                with open(CSV_PATH, "r", errors="ignore") as f:
                    sample = f.read(1024)
                    if sample.count(";") > sample.count(","):
                        delimiter = ";"
                    else:
                        delimiter = ","

                print(f"Utilizando delimitador: '{delimiter}'")

                # Tentar diferentes encodings
                encodings = ["utf-8", "latin1", "cp1252", "utf-8-sig"]
                for encoding in encodings:
                    try:
                        data = pd.read_csv(
                            CSV_PATH,
                            dtype=str,
                            skiprows=skip_rows,
                            sep=delimiter,
                            encoding=encoding,
                        )
                        print(f"Arquivo CSV carregado com encoding {encoding}")
                        break
                    except UnicodeDecodeError:
                        continue
                    except Exception as e:
                        raise e
                else:
                    raise Exception(
                        "Não foi possível determinar a codificação do arquivo CSV"
                    )

            except Exception as e:
                print(f"Erro ao processar CSV: {str(e)}")
                print("Tentando método padrão...")
                data = pd.read_csv(CSV_PATH, dtype=str)

        # Debug: Print actual column names to identify issues
        print("Colunas originais no arquivo:")
        print(data.columns.tolist())

        # Clean column names by stripping whitespace
        data.columns = data.columns.str.strip()

        # Debug: Print cleaned column names
        print("Colunas após limpeza:")
        print(data.columns.tolist())

        # Mapear e renomear colunas para o formato esperado
        column_mapping = {
            "CPF/CNPJ": "CPF_CNPJ",
            "SERVIÇO": "SERVICO",
            "REFERENCIA": "REFERENCIA",
            "VENCIMENTO": "VENCIMENTO",
            "VALOR": "VALOR",
            "NOTA FISCAL": "NF",  # Changed from "nf" to "NOTA FISCAL"
            "INFORMAÇÕES ADICIONAIS": "INFO_ADICIONAIS",
        }

        # Renomear colunas
        data = data.rename(columns=column_mapping)

        # Limpar dados - remover espaços extras e converter vírgula para ponto no valor
        data["VALOR"] = data["VALOR"].str.strip().str.replace(",", ".")
        data["VALOR"] = data["VALOR"].str.replace(" ", "")

        # Combinar NF com INFORMAÇÕES ADICIONAIS em um único campo
        data["INFO_COMBINADA"] = (
            "NF: " + data["NF"].fillna("") + " - " + data["INFO_ADICIONAIS"].fillna("")
        )

        print("Dados carregados do CSV:")
        print(data.head())

        # Processar cada linha
        for index, row in data.iterrows():
            dados = row.to_dict()
            print(
                f"\nProcessando: {dados['CPF_CNPJ']} - Ref. {dados['REFERENCIA']} - Valor: {dados['VALOR']}"
            )

            try:
                if not preencher_formulario(dados):
                    print("Interrupção solicitada durante preenchimento do formulário")
                    break

                if not baixar_pdf(
                    dados["CPF_CNPJ"],
                    dados["REFERENCIA"],
                    dados.get("INFO_ADICIONAIS", ""),
                    dados.get("VALOR", ""),
                ):
                    print("Falha na emissão")
            except Exception as e:
                print(f"Erro ao processar linha {index+1}: {str(e)}")

    except Exception as e:
        print(f"Erro ao processar o arquivo CSV: {str(e)}")
        print("Detalhes do erro:")
        import traceback

        traceback.print_exc()

    # Fechar o navegador ao finalizar
    close_browser()
