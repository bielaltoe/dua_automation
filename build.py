#!/usr/bin/env python3
"""
Script para compilar o DUA Automation em um executável
Este script utiliza PyInstaller para criar um executável autônomo do aplicativo
"""

import os
import sys
import shutil
import subprocess
import platform

# Configurações
APP_NAME = "DUA Automation"
VERSION = "1.0.1"  # Incrementado para nova versão no GitHub
MAIN_SCRIPT = "run_ui.py"

# Usar logo_new.png como ícone do executável gerado
LOGO_SOURCE = os.path.join("resources", "logo_new.png")
APP_ICON = os.path.join("resources", "app_icon_exe.ico")  # Para Windows
APP_ICNS = os.path.join("resources", "app_icon_exe.icns")  # Para macOS

# Adicionar a opção de criar a pasta .github se não existir
def ensure_github_files():
    """Cria a estrutura de arquivos .github se não existir"""
    github_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".github")
    templates_dir = os.path.join(github_dir, "ISSUE_TEMPLATE")
    
    # Criar diretórios se não existirem
    os.makedirs(templates_dir, exist_ok=True)
    
    print(f"Estrutura de diretórios GitHub verificada: {github_dir}")


def clean_build_dirs():
    """Limpa diretórios de build anteriores"""
    print("Limpando diretórios de build anteriores...")
    for dir_name in ["build", "dist"]:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"  - Diretório {dir_name} removido")


def check_dependencies():
    """Verifica se as dependências necessárias estão instaladas"""
    print("Verificando dependências...")

    # Lista de pacotes necessários para o build
    required_packages = [
        "PyInstaller",
        "pyqt6",
        "selenium",
        "webdriver-manager",
        "pandas",
        "numpy",
        "pydub",
        "SpeechRecognition",
        "aiohttp",
        "requests",
        "pillow",
        "psutil",
        "pathlib",
        "openpyxl",  # Adicionado para suporte a Excel
        "xlrd",      # Adicionado para suporte a Excel
    ]

    missing_packages = []
    for package in required_packages:
        try:
            __import__(package.lower().replace("-", "_"))
            print(f"  ✓ {package} está instalado")
        except ImportError:
            missing_packages.append(package)
            print(f"  ✗ {package} não encontrado")

    if missing_packages:
        print("\nPacotes ausentes encontrados. Instalando...")
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install"] + missing_packages
        )
        print("Dependências instaladas com sucesso!")
    
    # Verificar FFmpeg
    print("\nVerificando instalação do FFmpeg (obrigatório para processamento de áudio)...")
    try:
        result = subprocess.run(
            ["ffmpeg", "-version"], 
            stdout=subprocess.PIPE, 
            stderr=subprocess.PIPE,
            text=True,
            check=False
        )
        if result.returncode == 0:
            print("  ✓ FFmpeg está instalado")
        else:
            print("  ✗ FFmpeg não está instalado ou não está no PATH")
            print("\n⚠️ AVISO: FFmpeg é obrigatório para o funcionamento do sistema.")
            print("Por favor, instale o FFmpeg conforme as instruções no README.")
    except FileNotFoundError:
        print("  ✗ FFmpeg não encontrado. Este componente é OBRIGATÓRIO.")
        print("\n⚠️ AVISO: Sem o FFmpeg, a resolução de CAPTCHA não funcionará!")
        print("Por favor, instale o FFmpeg conforme as instruções no README.")


def convert_icon_if_needed():
    """Converte logo_new.png para o formato adequado para o executável"""
    system = platform.system()

    # Verificar se o arquivo fonte logo_new.png existe
    if not os.path.exists(LOGO_SOURCE):
        print(f"⚠️ Arquivo logo_new.png não encontrado em {LOGO_SOURCE}")
        print("  Usando ícones padrão no lugar...")
        return

    # Criar o diretório resources se não existir
    os.makedirs("resources", exist_ok=True)

    # Verifica se já temos os ícones necessários
    if system == "Windows" and not os.path.exists(APP_ICON):
        print(f"Convertendo {LOGO_SOURCE} para formato .ico...")
        # Tente converter de png para ico usando pillow
        try:
            from PIL import Image

            if os.path.exists(LOGO_SOURCE):
                img = Image.open(LOGO_SOURCE)
                # Redimensionar para tamanhos de ícones comuns do Windows
                icon_sizes = [
                    (16, 16),
                    (32, 32),
                    (48, 48),
                    (64, 64),
                    (128, 128),
                    (256, 256),
                ]
                img.save(APP_ICON, sizes=icon_sizes)
                print(f"  ✓ Logo convertido para ícone do executável: {APP_ICON}")
            else:
                print(f"  ✗ Arquivo de origem não encontrado: {LOGO_SOURCE}")
        except Exception as e:
            print(f"  ✗ Erro ao converter logo para ícone: {e}")

    elif system == "Darwin" and not os.path.exists(APP_ICNS):
        print(
            f"Conversão de {LOGO_SOURCE} para .icns requer ferramentas extras no macOS."
        )
        print("Por favor, converta manualmente o arquivo PNG para ICNS.")


def build_executable():
    """Constrói o executável usando PyInstaller"""
    print("\nIniciando build com PyInstaller...")

    system = platform.system()
    icon_param = []

    # Configurar parâmetro de ícone baseado no sistema operacional
    if system == "Windows" and os.path.exists(APP_ICON):
        icon_param = ["--icon", APP_ICON]
    elif system == "Darwin" and os.path.exists(APP_ICNS):
        icon_param = ["--icon", APP_ICNS]

    # Construir o comando PyInstaller
    pyinstaller_cmd = [
        "pyinstaller",
        "--name",
        APP_NAME.replace(" ", "_"),
        "--windowed",  # Sem console/terminal
        "--onefile",  # Arquivo único
        "--clean",  # Limpar cache
        "--add-data",
        f"resources{os.pathsep}resources",  # Incluir recursos
        "--noconfirm",  # Não confirmar sobrescrita
    ]

    # Adicionar parâmetro de ícone se disponível
    if icon_param:
        pyinstaller_cmd.extend(icon_param)

    # Adicionar o script principal
    pyinstaller_cmd.append(MAIN_SCRIPT)

    # Executar o PyInstaller
    try:
        subprocess.check_call(pyinstaller_cmd)
        print("\n✓ Build concluído com sucesso!")

        # Mostrar o local do executável
        if system == "Windows":
            exe_path = os.path.join("dist", f"{APP_NAME.replace(' ', '_')}.exe")
        elif system == "Darwin":
            exe_path = os.path.join("dist", f"{APP_NAME.replace(' ', '_')}.app")
        else:  # Linux
            exe_path = os.path.join("dist", APP_NAME.replace(" ", "_"))

        if os.path.exists(exe_path):
            print(f"\nExecutável gerado em: {os.path.abspath(exe_path)}")
        else:
            print(f"\nExecutável esperado em: {os.path.abspath(exe_path)}")
            print("Mas o arquivo não foi encontrado. Verifique erros no build.")

    except subprocess.CalledProcessError as e:
        print(f"\n✗ Erro durante o build: {e}")
        return False

    return True


def create_spec_file():
    """Cria um arquivo .spec personalizado para PyInstaller"""
    spec_content = f"""# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Listar todos os arquivos de recursos
resources = [
    ('resources/*', 'resources'),
    ('chrome-portable/*', 'chrome-portable') if os.path.exists('chrome-portable') else None,  # Include portable Chrome if it exists
]
# Remover entradas None
resources = [r for r in resources if r is not None]

a = Analysis(
    ['{MAIN_SCRIPT}'],
    pathex=[],
    binaries=[],
    datas=resources,
    hiddenimports=['PyQt6.sip', 'selenium'],
    hookspath=[],
    hooksconfig={{}},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='{APP_NAME.replace(" ", "_")}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['{APP_ICON if platform.system() == "Windows" else APP_ICNS}'] if os.path.exists('{APP_ICON if platform.system() == "Windows" else APP_ICNS}') else None,
)

# Para macOS, cria um .app bundle
if platform.system() == 'Darwin':
    app = BUNDLE(
        exe,
        name='{APP_NAME.replace(" ", "_")}.app',
        icon='{APP_ICNS}',
        bundle_identifier='com.duaautomation.app',
        info_plist={{
            'CFBundleShortVersionString': '{VERSION}',
            'CFBundleVersion': '{VERSION}',
            'NSHighResolutionCapable': 'True'
        }},
    )
"""

    with open("dua_automation.spec", "w") as f:
        f.write(spec_content)

    print("Arquivo .spec personalizado criado: dua_automation.spec")
    return "dua_automation.spec"


def main():
    print(f"==== Compilando {APP_NAME} v{VERSION} ====\n")

    # Verificar e criar estrutura GitHub
    ensure_github_files()

    # Verificar dependências
    check_dependencies()

    # Limpar builds anteriores
    clean_build_dirs()

    # Converter ícone se necessário
    convert_icon_if_needed()

    # Verificar Chrome portável
    print("\nVerificando Chrome portável para empacotar...")
    chrome_portable_dir = os.path.join(
        os.path.dirname(os.path.abspath(__file__)), "chrome-portable"
    )
    if not os.path.exists(chrome_portable_dir):
        download_chrome = input(
            "Chrome portável não encontrado. Deseja baixar agora? (s/n): "
        )
        if download_chrome.lower() == "s":
            try:
                print("Inicializando download do Chrome portável...")
                # Import module here to avoid circular imports
                import get_dua

                get_dua.download_portable_chrome()
            except Exception as e:
                print(f"Erro ao baixar Chrome portável: {str(e)}")
                print(
                    "Continuando sem Chrome portável. O executável usará o Chrome do sistema."
                )
    else:
        print(f"  ✓ Diretório do Chrome portável encontrado em: {chrome_portable_dir}")

    # Criar arquivo .spec personalizado
    spec_file = create_spec_file()

    # Incluir arquivos de documentação do GitHub
    print("\nVerificando arquivos de documentação para GitHub...")
    docs_files = [
        "README.md", 
        "LICENSE", 
        "CONTRIBUTING.md", 
        ".github/ISSUE_TEMPLATE/bug_report.md",
        ".github/ISSUE_TEMPLATE/feature_request.md",
        ".github/pull_request_template.md",
        ".gitignore"
    ]
    
    for doc_file in docs_files:
        if os.path.exists(doc_file):
            print(f"  ✓ Arquivo {doc_file} encontrado")
        else:
            print(f"  ! Arquivo {doc_file} não encontrado")

    # Executar PyInstaller com o arquivo .spec
    print("\nExecutando PyInstaller com arquivo .spec personalizado...")
    try:
        subprocess.check_call(["pyinstaller", "--clean", spec_file])
        print("\n✓ Build concluído com sucesso usando arquivo .spec!")

        # Mostrar o local do executável
        system = platform.system()
        if system == "Windows":
            exe_path = os.path.join("dist", f"{APP_NAME.replace(' ', '_')}.exe")
        elif system == "Darwin":
            exe_path = os.path.join("dist", f"{APP_NAME.replace(' ', '_')}.app")
        else:  # Linux
            exe_path = os.path.join("dist", APP_NAME.replace(" ", "_"))

        if os.path.exists(exe_path):
            print(f"\nExecutável gerado em: {os.path.abspath(exe_path)}")
        else:
            print(f"\nExecutável esperado em: {os.path.abspath(exe_path)}")
            print("Mas o arquivo não foi encontrado. Verifique erros no build.")
    except subprocess.CalledProcessError as e:
        print(f"\n✗ Erro durante o build com .spec: {e}")
        print("Tentando método alternativo de build...")

        # Tentar método alternativo
        build_executable()

    print("\nProcesso de build concluído!")


if __name__ == "__main__":
    main()
