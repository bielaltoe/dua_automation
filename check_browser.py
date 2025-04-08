"""
Utilitário para verificar instalação do Chrome e webdriver
"""

import os
import sys
import platform
import subprocess
import webbrowser
from pathlib import Path
import pkg_resources


def get_chrome_version():
    """Tenta obter a versão do Chrome instalada no sistema"""
    try:
        # Locais comuns do executável do Chrome por sistema operacional
        if platform.system() == "Windows":
            locations = [
                os.path.join(
                    os.environ.get("PROGRAMFILES", "C:\\Program Files"),
                    "Google\\Chrome\\Application\\chrome.exe",
                ),
                os.path.join(
                    os.environ.get("PROGRAMFILES(X86)", "C:\\Program Files (x86)"),
                    "Google\\Chrome\\Application\\chrome.exe",
                ),
                os.path.join(
                    os.environ.get("LOCALAPPDATA", ""),
                    "Google\\Chrome\\Application\\chrome.exe",
                ),
            ]

            for location in locations:
                if os.path.exists(location):
                    # No Windows, use wmic para obter a versão
                    try:
                        # Fix the f-string backslash issue by using double backslashes or raw string
                        version_result = subprocess.run(
                            [
                                "wmic",
                                "datafile",
                                "where",
                                'name="' + location.replace("\\", "\\\\") + '"',
                                "get",
                                "Version",
                                "/value",
                            ],
                            capture_output=True,
                            text=True,
                            check=True,
                        )
                        version_line = version_result.stdout.strip()
                        # Extrair a versão do formato "Version=xx.xx.xx.xx"
                        if "Version=" in version_line:
                            version = version_line.split("=")[1].strip()
                            return location, version
                    except:
                        pass

        elif platform.system() == "Darwin":  # macOS
            try:
                locations = [
                    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                    os.path.expanduser(
                        "~/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
                    ),
                ]

                for location in locations:
                    if os.path.exists(location):
                        version_result = subprocess.run(
                            [location, "--version"], capture_output=True, text=True
                        )
                        version_str = version_result.stdout.strip()
                        if "Chrome" in version_str:
                            version = version_str.split("Chrome")[1].strip()
                            return location, version
            except:
                pass

        else:  # Linux
            try:
                locations = [
                    "/usr/bin/google-chrome",
                    "/usr/bin/chromium-browser",
                    "/usr/bin/chromium",
                ]

                for location in locations:
                    if os.path.exists(location):
                        version_result = subprocess.run(
                            [location, "--version"], capture_output=True, text=True
                        )
                        version_str = version_result.stdout.strip()
                        if "Chrome" in version_str:
                            version = version_str.split("Chrome")[1].strip()
                            return location, version
            except:
                pass

        return None, "Não encontrado"
    except Exception as e:
        return None, f"Erro: {str(e)}"


def check_selenium_version():
    """Verifica a versão do Selenium instalada"""
    try:
        selenium_version = pkg_resources.get_distribution("selenium").version
        return selenium_version
    except:
        return "Não instalado"


def check_webdriver_manager():
    """Verifica se webdriver-manager está instalado"""
    try:
        wdm_version = pkg_resources.get_distribution("webdriver-manager").version
        return wdm_version
    except:
        return "Não instalado"


def open_chrome_download_page():
    """Abre a página de download do Chrome"""
    url = "https://www.google.com/chrome/"
    webbrowser.open(url)


def check_system():
    """Verifica o ambiente do sistema"""
    chrome_path, chrome_version = get_chrome_version()
    selenium_version = check_selenium_version()
    wdm_version = check_webdriver_manager()

    print("\n=== Verificação do Ambiente para DUA Automation ===\n")
    print(
        f"Sistema Operacional: {platform.system()} {platform.release()} ({platform.architecture()[0]})"
    )
    print(f"Python: {platform.python_version()} ({sys.executable})")
    print(f"Selenium: {selenium_version}")
    print(f"WebDriver Manager: {wdm_version}")

    if chrome_path:
        print(f"Google Chrome: {chrome_version}")
        print(f"  Localização: {chrome_path}")
        return True
    else:
        print("\n❌ Google Chrome NÃO ENCONTRADO!")
        print("\nO DUA Automation requer o Google Chrome para funcionar.")
        print("Por favor, instale o Chrome e tente novamente.")

        choice = input("\nDeseja abrir a página de download do Chrome? (s/n): ")
        if choice.lower() == "s":
            open_chrome_download_page()

        return False


if __name__ == "__main__":
    if check_system():
        print("\n✅ Ambiente verificado. Chrome encontrado.")
        print("O sistema está pronto para executar o DUA Automation.")
    else:
        print(
            "\n❌ Verificação falhou. Por favor, instale o Chrome antes de usar o DUA Automation."
        )

    input("\nPressione Enter para sair...")
