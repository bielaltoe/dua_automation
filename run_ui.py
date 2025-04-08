#!/usr/bin/env python3
import sys
import os
import traceback
from PyQt6.QtCore import QLibraryInfo
from PyQt6.QtWidgets import QMessageBox, QApplication

# Enable debugging
os.environ["DEBUG"] = "1"

# Determine if running as executable or in development
is_frozen = getattr(sys, "frozen", False)

# Set up proper paths for resources and output folders
if is_frozen:
    # Running as executable
    exe_dir = os.path.dirname(sys.executable)

    # Set PDF directory to a sibling folder to the executable
    pdf_dir = os.path.join(exe_dir, "pdfs_gerados")
    os.environ["PDF_DIR"] = pdf_dir
    os.makedirs(pdf_dir, exist_ok=True)

    # Path to resources when frozen
    resources_dir = os.path.join(exe_dir, "resources")
    if not os.path.exists(resources_dir):
        try:
            os.makedirs(resources_dir)
            print(f"Created resources directory: {resources_dir}")
        except Exception as e:
            print(f"Could not create resources directory: {e}")
else:
    # Running in development mode
    script_dir = os.path.dirname(os.path.abspath(__file__))

    # Set PDF directory
    pdf_dir = os.path.join(script_dir, "pdfs_gerados")
    os.environ["PDF_DIR"] = pdf_dir
    os.makedirs(pdf_dir, exist_ok=True)

    # Resources directory
    resources_dir = os.path.join(script_dir, "resources")
    if not os.path.exists(resources_dir):
        try:
            os.makedirs(resources_dir)
            print(f"Created resources directory: {resources_dir}")
        except Exception as e:
            print(f"Could not create resources directory: {e}")

print(f"PDFs will be saved to: {pdf_dir}")
print(f"Resources directory: {resources_dir}")
print(f"Running as executable: {is_frozen}")


def check_chrome_installed():
    """Verifica se o Chrome está instalado antes de prosseguir"""
    try:
        from check_browser import get_chrome_version

        chrome_path, chrome_version = get_chrome_version()

        if not chrome_path:
            # Importar PyQt para mostrar mensagem
            app = QApplication(sys.argv)

            msg = QMessageBox()
            msg.setIcon(QMessageBox.Icon.Critical)
            msg.setWindowTitle("Erro - Chrome não encontrado")
            msg.setText("O Google Chrome não foi encontrado no seu sistema.")
            msg.setInformativeText(
                "O DUA Automation requer o Google Chrome para funcionar.\n\n"
                "Por favor, instale o Chrome e tente novamente.\n\n"
                "Deseja abrir a página de download do Chrome?"
            )
            msg.setStandardButtons(
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )

            if msg.exec() == QMessageBox.StandardButton.Yes:
                import webbrowser

                webbrowser.open("https://www.google.com/chrome/")

            return False

        print(f"Chrome encontrado: versão {chrome_version}")
        return True
    except Exception as e:
        print(f"Erro ao verificar Chrome: {e}")
        # Continuar mesmo com erro na verificação
        return True


if __name__ == "__main__":
    # Show version information
    print(f"Python version: {sys.version}")
    print(f"Running from: {os.getcwd()}")

    # Try to ensure xcb plugin is available by setting paths
    os.environ["QT_DEBUG_PLUGINS"] = "1"  # Enable plugin debugging

    # Verificar instalação do Chrome
    if not check_chrome_installed():
        print("Aplicação terminada: Chrome não encontrado.")
        sys.exit(1)

    # Create fallback mechanism in case xcb is not available
    try:
        # Import early to catch any problems
        print("Importing UI modules...")
        from ui import QApplication, DUAAutomationUI, LogMessage

        print("UI modules imported successfully")

        # Basic exception handler to show errors
        def excepthook(exc_type, exc_value, exc_tb):
            error_message = "".join(
                traceback.format_exception(exc_type, exc_value, exc_tb)
            )
            print(f"ERROR: {error_message}")

        sys.excepthook = excepthook

        print("Starting application...")
        app = QApplication(sys.argv)
        app.setStyle("Fusion")  # Modern look across platforms
        window = DUAAutomationUI()
        window.show()
        print("Application started successfully")

        # Add a startup message to the log
        window.log_text.append_log(LogMessage("Aplicação iniciada com sucesso.", 1))

        sys.exit(app.exec())
    except Exception as e:
        print(f"Error starting Qt application: {e}")
        traceback.print_exc()

        # Try with specific platform if xcb fails
        if "Could not load the Qt platform plugin" in str(e):
            print("Trying fallback to a different platform plugin...")

            # Try wayland, then offscreen as fallbacks
            for platform in ["wayland", "offscreen"]:
                try:
                    os.environ["QT_QPA_PLATFORM"] = platform
                    print(f"Trying with {platform} platform...")

                    # Re-import inside the loop since we're changing the environment
                    from ui import QApplication, DUAAutomationUI

                    app = QApplication(sys.argv)
                    app.setStyle("Fusion")
                    window = DUAAutomationUI()
                    window.show()
                    sys.exit(app.exec())
                except Exception as fallback_error:
                    print(f"Failed with {platform}: {fallback_error}")
                    traceback.print_exc()

            # If we get here, none of the fallbacks worked
            print("\nERROR: Could not initialize any Qt platform plugin.")
            print(
                "Please install the required dependencies with one of these commands:"
            )
            print("  Ubuntu/Debian: sudo apt-get install libxcb-cursor0")
            print("  Fedora: sudo dnf install libxcb-cursor")
            print("  Arch Linux: sudo pacman -S xcb-util-cursor")
