import sys
import os
import pandas as pd
import threading
import queue
import time
import datetime
from pathlib import Path

from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QPushButton,
    QLabel,
    QFileDialog,
    QTableView,
    QHeaderView,
    QTextEdit,
    QProgressBar,
    QTabWidget,
    QLineEdit,
    QFormLayout,
    QMessageBox,
    QGroupBox,
    QStatusBar,
    QStyle,
    QCheckBox,
    QComboBox,
    QSplitter,
    QDialog,
)
from PyQt6.QtCore import (
    Qt,
    QAbstractTableModel,
    QModelIndex,
    pyqtSignal,
    pyqtSlot,
    QThread,
    QSettings,
    QSize,
    QTimer,
)
from PyQt6.QtGui import (
    QFont,
    QIcon,
    QAction,
    QColor,
    QTextCharFormat,
    QTextCursor,
    QPixmap,
)

# Import functions from get_dua.py
from get_dua import (
    preencher_formulario,
    baixar_pdf,
    SERVICO_MAPPING,
    set_captcha_callback,
    captcha_solved_signal,
)

# Import the new captcha dialog
from captcha_dialog import CaptchaDialog


class DataFrameModel(QAbstractTableModel):
    """Model for displaying pandas DataFrame in a QTableView"""

    def __init__(self, data=None):
        super().__init__()
        self._data = data if data is not None else pd.DataFrame()

    def rowCount(self, parent=None):
        return len(self._data)

    def columnCount(self, parent=None):
        return len(self._data.columns)

    def data(self, index, role=Qt.ItemDataRole.DisplayRole):
        if not index.isValid() or role != Qt.ItemDataRole.DisplayRole:
            return None
        return str(self._data.iloc[index.row(), index.column()])

    def headerData(self, section, orientation, role=Qt.ItemDataRole.DisplayRole):
        if (
            orientation == Qt.Orientation.Horizontal
            and role == Qt.ItemDataRole.DisplayRole
        ):
            return str(self._data.columns[section])
        if (
            orientation == Qt.Orientation.Vertical
            and role == Qt.ItemDataRole.DisplayRole
        ):
            return str(section + 1)
        return None

    def update_data(self, data):
        self.beginResetModel()
        self._data = data
        self.endResetModel()


class LogMessage:
    """Class to represent a log message with severity level"""

    INFO = 0
    SUCCESS = 1
    WARNING = 2
    ERROR = 3

    def __init__(self, text, level=INFO, timestamp=None):
        self.text = text
        self.level = level
        self.timestamp = timestamp or datetime.datetime.now()

    def format_timestamp(self):
        return self.timestamp.strftime("%H:%M:%S")

    def get_color(self):
        if self.level == LogMessage.SUCCESS:
            return QColor(0, 128, 0)  # Green
        elif self.level == LogMessage.WARNING:
            return QColor(255, 165, 0)  # Orange
        elif self.level == LogMessage.ERROR:
            return QColor(255, 0, 0)  # Red
        else:
            return QColor(0, 0, 0)  # Black for INFO


class EnhancedTextEdit(QTextEdit):
    """Enhanced QTextEdit with better log display capabilities"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setReadOnly(True)
        self.setFont(QFont("Consolas", 10))
        self.setLineWrapMode(QTextEdit.LineWrapMode.NoWrap)

    def append_log(self, message):
        # Create text format with the appropriate color
        text_format = QTextCharFormat()
        text_format.setForeground(message.get_color())

        # Make errors bold
        if message.level == LogMessage.ERROR:
            text_format.setFontWeight(QFont.Weight.Bold)

        # Add the formatted text
        cursor = self.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.End)

        # Add timestamp
        timestamp_format = QTextCharFormat()
        timestamp_format.setForeground(QColor(100, 100, 100))
        cursor.insertText(f"[{message.format_timestamp()}] ", timestamp_format)

        # Add the message
        cursor.insertText(f"{message.text}\n", text_format)

        # Ensure visible
        self.setTextCursor(cursor)
        self.ensureCursorVisible()


class LogHandler(QThread):
    """Thread for handling log messages"""

    log_signal = pyqtSignal(LogMessage)

    def __init__(self, message_queue):
        super().__init__()
        self.message_queue = message_queue
        self.running = True

    def run(self):
        while self.running:
            try:
                message = self.message_queue.get(timeout=0.1)
                self.log_signal.emit(message)
            except queue.Empty:
                continue

    def stop(self):
        self.running = False


class WorkerThread(QThread):
    """Thread for running the DUA automation process"""

    progress_signal = pyqtSignal(int, int)  # current, total
    finished_signal = pyqtSignal(bool)  # success/failure
    log_signal = pyqtSignal(LogMessage)
    status_signal = pyqtSignal(str, int)  # message, level
    captcha_signal = pyqtSignal()  # Signal for manual CAPTCHA intervention

    def __init__(self, data, pdf_dir):
        super().__init__()
        self.data = data
        self.pdf_dir = pdf_dir
        self.running = True
        self.total_success = 0
        self.total_failure = 0

    def stop(self):
        """Stop the worker thread safely"""
        self.running = False
        print("Stop flag set - thread will terminate at next check point")

    def run(self):
        total_rows = len(self.data)

        # Setup for logging
        import builtins
        import sys

        # Fix for executables: safely get stdout/stderr or create a fallback
        try:
            original_stdout = sys.stdout
            original_stderr = sys.stderr
        except:
            # Create a fallback stdout/stderr for executable environments
            class NullStream:
                def write(self, *args, **kwargs):
                    pass

                def flush(self):
                    pass

            original_stdout = NullStream()
            original_stderr = NullStream()

        # Direct log function that uses Qt signals
        def direct_log(message, level=LogMessage.INFO):
            # Check running flag at each log point
            if not self.running:
                return

            # Emit signal directly to update UI
            self.log_signal.emit(LogMessage(message, level))

            # Safely print to original stdout for terminal logging
            try:
                if original_stdout is not None:
                    original_stdout.write(f"{message}\n")
                    original_stdout.flush()
            except Exception:
                # Silently ignore errors when stdout is unavailable
                pass

        # Create a custom print function
        def custom_print(*args, **kwargs):
            # Check running flag at each print
            if not self.running:
                return

            message = " ".join(map(str, args))

            # Determine message level based on content
            level = LogMessage.INFO
            if (
                "erro" in message.lower()
                or "falha" in message.lower()
                or "error" in message.lower()
            ):
                level = LogMessage.ERROR
            elif (
                "sucesso" in message.lower()
                or "✅" in message
                or "concluído" in message.lower()
                or "resolvido" in message.lower()
                or "gerado e salvo" in message.lower()
            ):
                level = LogMessage.SUCCESS
            elif (
                "aguardando" in message.lower()
                or "tentando" in message.lower()
                or "encontrado" in message.lower()
                or "captcha" in message.lower()
            ):
                level = LogMessage.WARNING

            # Use direct_log to update UI
            direct_log(message, level)

        # Install our custom print function
        builtins.print = custom_print
        builtins.log_ui = direct_log  # Add a direct UI logger

        # Inform UI we're starting
        direct_log("🔄 Iniciando processamento de DUAs", LogMessage.INFO)
        direct_log("🛠️ Preparando ambiente e configurando sistema...", LogMessage.INFO)

        try:
            self.status_signal.emit("Iniciando processamento...", LogMessage.INFO)

            # Import get_dua modules and initialize browser
            direct_log(
                "🔄 Carregando módulos e inicializando navegador...", LogMessage.INFO
            )
            from get_dua import initialize_driver, close_browser

            # Registrar o callback para resolução manual de CAPTCHA
            set_captcha_callback(self.request_manual_captcha)

            driver = initialize_driver()
            direct_log("✅ Navegador iniciado com sucesso", LogMessage.SUCCESS)

            # Process each row
            for index, row in self.data.iterrows():
                # Check running flag at the start of each row processing
                if not self.running:
                    self.status_signal.emit(
                        "Processamento interrompido pelo usuário.", LogMessage.WARNING
                    )
                    direct_log(
                        "⚠️ Processamento interrompido pelo usuário", LogMessage.WARNING
                    )
                    break

                dados = row.to_dict()

                # Update status with current item
                status_msg = (
                    f"Processando: {dados['CPF_CNPJ']} - Ref. {dados['REFERENCIA']}"
                )
                self.status_signal.emit(status_msg, LogMessage.INFO)

                # Use both print and direct UI log (belt and suspenders)
                msg = f"📝 Processando item {index+1}/{total_rows}: CPF/CNPJ: {dados['CPF_CNPJ']} - Ref: {dados['REFERENCIA']} - Valor: R$ {dados['VALOR']}"
                direct_log(msg, LogMessage.INFO)

                try:
                    # Step 1: Preencher formulário
                    direct_log("🔄 Preenchendo formulário DUA...", LogMessage.INFO)
                    from get_dua import preencher_formulario

                    preencher_formulario(dados)

                    # Step 2: Baixar PDF
                    direct_log("🔄 Gerando e baixando o PDF...", LogMessage.INFO)
                    from get_dua import baixar_pdf

                    # Atualizado: Passar todos os parâmetros relevantes
                    success = baixar_pdf(
                        dados["CPF_CNPJ"],
                        dados["REFERENCIA"],
                        dados.get("INFO_ADICIONAIS", ""),
                        dados.get("VALOR", ""),
                    )

                    if success:
                        self.total_success += 1
                        direct_log(
                            f"✅ DUA gerado com sucesso para {dados['CPF_CNPJ']} - Ref: {dados['REFERENCIA']}",
                            LogMessage.SUCCESS,
                        )
                    else:
                        self.total_failure += 1
                        direct_log(
                            f"❌ Falha na emissão do DUA para {dados['CPF_CNPJ']} - Ref: {dados['REFERENCIA']}",
                            LogMessage.ERROR,
                        )

                except Exception as e:
                    self.total_failure += 1
                    direct_log(
                        f"❌ Erro ao processar item {index+1}: {str(e)}",
                        LogMessage.ERROR,
                    )

                # Update progress
                self.progress_signal.emit(index + 1, total_rows)

                # Add a small separator in the log
                direct_log(
                    "──────────────────────────────────────────────", LogMessage.INFO
                )

            # Final status update
            if self.total_success == total_rows:
                summary = f"🎉 Processamento concluído com sucesso! Total: {total_rows} DUA(s) gerado(s)."
                level = LogMessage.SUCCESS
            else:
                summary = f"⚠️ Processamento concluído. Sucessos: {self.total_success}, Falhas: {self.total_failure}"
                level = LogMessage.WARNING

            self.status_signal.emit(summary, level)
            direct_log(summary, level)

        except Exception as e:
            error_msg = f"❌ Erro crítico no processamento: {str(e)}"
            self.status_signal.emit(error_msg, LogMessage.ERROR)
            direct_log(error_msg, LogMessage.ERROR)

            import traceback

            error_details = traceback.format_exc()
            direct_log("Detalhes do erro:", LogMessage.ERROR)
            direct_log(error_details, LogMessage.ERROR)

        finally:
            # Restore stdout/stderr
            builtins.print = original_stdout

            try:
                direct_log("🔄 Finalizando navegador...", LogMessage.INFO)
                from get_dua import close_browser

                close_browser()
                direct_log("✅ Navegador finalizado com sucesso", LogMessage.SUCCESS)
            except:
                direct_log(
                    "⚠️ Não foi possível finalizar o navegador corretamente",
                    LogMessage.WARNING,
                )

            self.finished_signal.emit(self.total_failure == 0)
            direct_log("📊 Relatório final:", LogMessage.INFO)
            direct_log(f"   Total de DUAs processados: {total_rows}", LogMessage.INFO)
            direct_log(f"   Sucessos: {self.total_success}", LogMessage.SUCCESS)
            direct_log(
                f"   Falhas: {self.total_failure}",
                LogMessage.ERROR if self.total_failure > 0 else LogMessage.INFO,
            )

    def request_manual_captcha(self):
        """Método chamado quando o sistema precisa de intervenção manual para o CAPTCHA"""
        print("Thread de trabalho solicitando intervenção manual para CAPTCHA")
        self.captcha_signal.emit()


class DUAAutomationUI(QMainWindow):
    """Main UI window for DUA automation"""

    def __init__(self):
        super().__init__()
        self.settings = QSettings("DUA_Automation", "Settings")
        self.worker = None

        self.initUI()
        self.loadSettings()

    def initUI(self):
        # Definir ícone da aplicação (ícone pequeno para a barra de título)
        icon_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "resources", "app_icon.png"
        )

        # Definir logo para o cabeçalho (imagem maior)
        logo_path = os.path.join(
            os.path.dirname(os.path.abspath(__file__)), "resources", "company_logo.png"
        )

        # Tentar caminhos alternativos se os arquivos não existirem
        if not os.path.exists(icon_path):
            # Verificar outros ícones comuns que podem estar disponíveis
            alternative_icons = [
                "icon.png",
                "dua_icon.png",
                "favicon.png",
                "window_icon.png",
            ]

            for alt_icon in alternative_icons:
                alt_path = os.path.join(
                    os.path.dirname(os.path.abspath(__file__)), "resources", alt_icon
                )
                if os.path.exists(alt_path):
                    icon_path = alt_path
                    break

        # Verificar alternativas para logo se não existir
        if not os.path.exists(logo_path):
            alternative_logos = [
                "logo.png",
                "header_logo.png",
                "brand_logo.png",
                "company_brand.png",
            ]

            for alt_logo in alternative_logos:
                alt_path = os.path.join(
                    os.path.dirname(os.path.abspath(__file__)), "resources", alt_logo
                )
                if os.path.exists(alt_path):
                    logo_path = alt_path
                    break

        # Definir o ícone da janela se o arquivo existir
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.setWindowTitle("Emissão Automática de DUAs")
        self.setGeometry(100, 100, 1280, 900)  # Increased window size

        # Create central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Criar um cabeçalho atrativo com tema escuro
        header_widget = QWidget()
        header_widget.setStyleSheet(
            """
            background-color: #1e2130; 
            border-bottom: 1px solid #2d3446;
            padding: 15px;
            margin-bottom: 10px;
        """
        )
        header_layout = QHBoxLayout(header_widget)
        header_layout.setContentsMargins(20, 10, 20, 10)

        # Adicionar a logo no cabeçalho (usar logo_path)
        if os.path.exists(logo_path):
            logo_label = QLabel()
            logo_pixmap = QPixmap(logo_path).scaled(
                160,
                80,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation,
            )
            logo_label.setPixmap(logo_pixmap)
            logo_label.setAlignment(
                Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
            )
            header_layout.addWidget(logo_label)

        # Adicionar título e subtítulo
        title_container = QVBoxLayout()

        title_label = QLabel("Emissão Automática de DUAs")
        title_label.setStyleSheet("font-size: 24pt; font-weight: bold; color: #6dd5ed;")
        title_label.setAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignBottom
        )

        subtitle_label = QLabel("Sistema de Geração Automática de DUAs")
        subtitle_label.setStyleSheet(
            "font-size: 12pt; color: #a9b7c6; margin-top: -5px;"
        )
        subtitle_label.setAlignment(
            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop
        )

        title_container.addWidget(title_label)
        title_container.addWidget(subtitle_label)
        header_layout.addLayout(title_container)

        # Adicionar informações à direita (apenas data)
        info_container = QVBoxLayout()

        date_label = QLabel(datetime.datetime.now().strftime("%d/%m/%Y"))
        date_label.setStyleSheet("font-size: 10pt; color: #b8c6d1; text-align: right;")
        date_label.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop)

        info_container.addWidget(date_label)

        header_layout.addStretch(1)
        header_layout.addLayout(info_container)

        # Adicionar o cabeçalho ao layout principal
        main_layout.addWidget(header_widget)

        # Create tab widget
        tab_widget = QTabWidget()
        main_layout.addWidget(tab_widget)

        # Main tab
        main_tab = QWidget()
        # Use horizontal layout instead of vertical for the main tab
        main_tab_layout = QHBoxLayout(main_tab)
        tab_widget.addTab(main_tab, "Principal")

        # Create a horizontal splitter for the two-column layout
        main_splitter = QSplitter(Qt.Orientation.Horizontal)
        main_splitter.setChildrenCollapsible(False)  # Don't allow collapsing sections
        main_tab_layout.addWidget(main_splitter)

        # Create left panel for controls
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 10, 0)
        main_splitter.addWidget(left_panel)

        # Create right panel for table and logs
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(10, 0, 0, 0)
        main_splitter.addWidget(right_panel)

        # ========== LEFT PANEL CONTENT ==========

        # File selection area
        file_group = QGroupBox("Arquivo CSV")
        file_layout = QHBoxLayout(file_group)
        self.file_path_label = QLineEdit()
        self.file_path_label.setReadOnly(True)
        self.file_path_label.setPlaceholderText("Nenhum arquivo selecionado")
        file_layout.addWidget(self.file_path_label)

        self.select_file_btn = QPushButton("Selecionar Arquivo")
        self.select_file_btn.setToolTip(
            "Selecione um arquivo CSV ou Excel contendo os dados para geração de DUAs"
        )
        self.select_file_btn.clicked.connect(self.select_csv_file)
        file_layout.addWidget(self.select_file_btn)

        left_layout.addWidget(file_group)

        # Action buttons and status
        actions_group = QGroupBox("Ações")
        actions_layout = QVBoxLayout(actions_group)

        # Status indicator
        status_layout = QHBoxLayout()
        self.status_label = QLabel("Status: Pronto")
        self.status_label.setStyleSheet("font-weight: bold;")
        status_layout.addWidget(self.status_label)

        # Spacer and counters
        status_layout.addStretch()
        self.success_count = QLabel("Sucessos: 0")
        self.success_count.setStyleSheet("color: green; font-weight: bold;")
        status_layout.addWidget(self.success_count)

        self.failure_count = QLabel("Falhas: 0")
        self.failure_count.setStyleSheet("color: red; font-weight: bold;")
        status_layout.addWidget(self.failure_count)

        actions_layout.addLayout(status_layout)

        # Buttons
        buttons_layout = QVBoxLayout()
        self.start_btn = QPushButton("Iniciar Processamento")
        self.start_btn.setToolTip(
            "Inicia o processamento automático dos DUAs com base nos dados carregados"
        )
        self.start_btn.setMinimumHeight(40)  # Make button taller
        self.start_btn.clicked.connect(self.start_processing)
        self.start_btn.setEnabled(False)
        buttons_layout.addWidget(self.start_btn)

        self.stop_btn = QPushButton("Parar")
        self.stop_btn.setToolTip("Interrompe o processamento atual")
        self.stop_btn.setMinimumHeight(40)  # Make button taller
        self.stop_btn.clicked.connect(self.stop_processing)
        self.stop_btn.setEnabled(False)
        buttons_layout.addWidget(self.stop_btn)

        actions_layout.addLayout(buttons_layout)
        left_layout.addWidget(actions_group)

        # Progress area
        progress_group = QGroupBox("Progresso")
        progress_layout = QVBoxLayout(progress_group)
        self.progress_bar = QProgressBar()
        self.progress_bar.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.progress_bar.setFormat("%v/%m - %p%")
        progress_layout.addWidget(self.progress_bar)

        left_layout.addWidget(progress_group)

        # Add stretch to push everything to the top in the left panel
        left_layout.addStretch()

        # ========== RIGHT PANEL CONTENT ==========

        # Create vertical splitter for table and logs in right panel
        right_splitter = QSplitter(Qt.Orientation.Vertical)
        right_splitter.setChildrenCollapsible(False)
        right_layout.addWidget(right_splitter)

        # Table view for CSV data
        table_group = QGroupBox("Dados do Arquivo")
        table_layout = QVBoxLayout(table_group)
        self.table_view = QTableView()
        self.table_view.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.ResizeToContents
        )
        self.table_view.setMinimumHeight(250)  # Ensure minimum height
        self.table_model = DataFrameModel()
        self.table_view.setModel(self.table_model)
        table_layout.addWidget(self.table_view)
        right_splitter.addWidget(table_group)

        # Log area with filter
        log_group = QGroupBox("Log")
        log_layout = QVBoxLayout(log_group)

        # Log filter
        filter_layout = QHBoxLayout()
        filter_layout.addWidget(QLabel("Exibir:"))
        self.log_filter = QComboBox()
        self.log_filter.addItems(
            [
                "Todos os logs",
                "Apenas informações",
                "Apenas sucessos",
                "Apenas avisos",
                "Apenas erros",
            ]
        )
        self.log_filter.setToolTip("Filtra as mensagens de log por tipo")
        self.log_filter.currentIndexChanged.connect(self.apply_log_filter)
        filter_layout.addWidget(self.log_filter)
        filter_layout.addStretch()

        # Clear log button
        self.clear_log_btn = QPushButton("Limpar Log")
        self.clear_log_btn.setToolTip("Limpa todas as mensagens de log")
        self.clear_log_btn.clicked.connect(self.clear_log)
        filter_layout.addWidget(self.clear_log_btn)

        log_layout.addLayout(filter_layout)

        # Enhanced log text area
        self.log_text = EnhancedTextEdit()
        self.log_text.setMinimumHeight(250)  # Ensure minimum height
        log_layout.addWidget(self.log_text)
        right_splitter.addWidget(log_group)

        # Set initial sizes for both splitters
        main_splitter.setSizes([300, 900])  # Left panel smaller, right panel larger
        right_splitter.setSizes([500, 500])  # Equal space for table and logs initially

        # Settings tab
        settings_tab = QWidget()
        settings_layout = QVBoxLayout(settings_tab)
        tab_widget.addTab(settings_tab, "Configurações")

        # PDF directory setting
        pdf_group = QGroupBox("Diretório para PDFs")
        pdf_layout = QHBoxLayout(pdf_group)
        self.pdf_dir_edit = QLineEdit()
        self.pdf_dir_edit.setText(os.path.abspath("pdfs_gerados"))
        self.pdf_dir_edit.setToolTip("Diretório onde os PDFs dos DUAs serão salvos")
        pdf_layout.addWidget(self.pdf_dir_edit)

        self.select_pdf_dir_btn = QPushButton("Selecionar Diretório")
        self.select_pdf_dir_btn.setToolTip(
            "Escolha o diretório onde os PDFs serão salvos"
        )
        self.select_pdf_dir_btn.clicked.connect(self.select_pdf_directory)
        pdf_layout.addWidget(self.select_pdf_dir_btn)

        settings_layout.addWidget(pdf_group)

        # Add help/instructions tab
        help_tab = QWidget()
        help_layout = QVBoxLayout(help_tab)
        tab_widget.addTab(help_tab, "Instruções")

        help_text = EnhancedTextEdit()
        help_text.setReadOnly(True)
        # Instructions content with formatting
        instructions = """<h2 style="color: #2a5885;">Instruções de Uso da Emissão Automática de DUAs</h2>
        
<h3 style="color: #366998;">Visão Geral</h3>
<p>Este sistema automatiza a emissão de DUAs (Documento Único de Arrecadação) para o estado do Espírito Santo, eliminando a necessidade de preenchimento manual do formulário na página da SEFAZ-ES.</p>
        
<h3 style="color: #366998;">Passo a Passo</h3>
<ol>
  <li><b>Preparar arquivo CSV ou Excel</b> - Crie um arquivo com as seguintes colunas:
    <ul>
      <li>CPF/CNPJ: número do contribuinte</li>
      <li>SERVIÇO: código do serviço (ex: 138-4)</li>
      <li>REFERENCIA: mês/ano de referência (ex: 01/2024)</li>
      <li>VENCIMENTO: data de vencimento (ex: 10/01/2024)</li>
      <li>VALOR: valor do DUA (ex: 123.45)</li>
      <li>NOTA FISCAL: número da nota fiscal (opcional)</li>
      <li>INFORMAÇÕES ADICIONAIS: texto adicional (opcional)</li>
    </ul>
  </li> 
  <li><b>Selecionar arquivo</b> - Clique em "Selecionar Arquivo" e escolha o arquivo preparado</li>
  <li><b>Verificar dados</b> - Confira se os dados foram carregados corretamente na tabela</li>
  <li><b>Configurar diretório</b> - Na aba "Configurações", defina onde os PDFs serão salvos</li>
  <li><b>Iniciar processamento</b> - Clique em "Iniciar Processamento" para começar</li>
  <li><b>Acompanhar progresso</b> - O sistema processará cada linha do arquivo e salvará os PDFs no diretório configurado</li>
</ol>
        
<h3 style="color: #366998;">Solução de problemas</h3>
<ul>
  <li><b>CAPTCHAs</b>: O sistema tentará resolver automaticamente os CAPTCHAs. Em caso de falha, você precisará resolvê-los manualmente.</li>
  <li><b>Falhas de conexão</b>: Verifique sua conexão com a internet.</li>
  <li><b>Erros no arquivo</b>: Certifique-se de que seu arquivo está no formato correto e contém todas as colunas necessárias.</li>
</ul>
        
<h3 style="color: #366998;">Observações importantes</h3>
<ul>
  <li>Mantenha o navegador visível durante o processamento para acompanhar o progresso.</li>
  <li>Não interaja com o navegador durante o processamento automático.</li>
  <li>Os PDFs gerados ficarão disponíveis no diretório configurado.</li>
  <li>Consulte o log para detalhes sobre qualquer erro que ocorrer.</li>
</ul>
"""
        help_text.setHtml(instructions)
        help_layout.addWidget(help_text)

        # Add stretch to push everything to the top
        settings_layout.addStretch()

        # Status bar
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.statusBar.showMessage("Pronto")

        # Initialize log with a welcome message
        self.log_text.append_log(
            LogMessage(
                "🚀 Emissão Automática de DUAs inicializado. Selecione um arquivo para começar.",
                LogMessage.INFO,
            )
        )
        self.log_text.append_log(
            LogMessage(
                "ℹ️ Vá para a aba 'Instruções' para obter ajuda sobre como usar o sistema.",
                LogMessage.INFO,
            )
        )

    def loadSettings(self):
        pdf_dir = self.settings.value("pdf_directory")
        if pdf_dir:
            self.pdf_dir_edit.setText(pdf_dir)

    def saveSettings(self):
        self.settings.setValue("pdf_directory", self.pdf_dir_edit.text())

    def apply_log_filter(self, index):
        # Implement log filtering functionality
        try:
            # Save current cursor position
            cursor_pos = self.log_text.textCursor().position()

            # Get all log messages (we'll need to iterate through the document)
            self.log_text.clear()

            # Apply filter based on the selected index
            if index == 0:  # "Todos os logs"
                self.log_text.append_log(
                    LogMessage("🔍 Mostrando todos os logs", LogMessage.INFO)
                )
            elif index == 1:  # "Apenas informações"
                self.log_text.append_log(
                    LogMessage(
                        "🔍 Filtrando: apenas mensagens informativas", LogMessage.INFO
                    )
                )
            elif index == 2:  # "Apenas sucessos"
                self.log_text.append_log(
                    LogMessage(
                        "🔍 Filtrando: apenas mensagens de sucesso", LogMessage.SUCCESS
                    )
                )
            elif index == 3:  # "Apenas avisos"
                self.log_text.append_log(
                    LogMessage("🔍 Filtrando: apenas avisos", LogMessage.WARNING)
                )
            elif index == 4:  # "Apenas erros"
                self.log_text.append_log(
                    LogMessage("🔍 Filtrando: apenas erros", LogMessage.ERROR)
                )

        except Exception as e:
            self.log_text.append_log(
                LogMessage(f"Erro ao aplicar filtro: {str(e)}", LogMessage.ERROR)
            )

    def clear_log(self):
        self.log_text.clear()
        self.log_text.append_log(LogMessage("Log limpo pelo usuário", LogMessage.INFO))

    def select_csv_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Selecionar Arquivo",
            "",
            "Arquivos de Dados (*.csv *.xlsx *.xls);;Arquivos CSV (*.csv);;Arquivos Excel (*.xlsx *.xls);;Todos os Arquivos (*)",
        )
        if file_path:
            self.file_path_label.setText(file_path)
            self.load_csv_data(file_path)

    def select_pdf_directory(self):
        dir_path = QFileDialog.getExistingDirectory(
            self, "Selecionar Diretório para PDFs", self.pdf_dir_edit.text()
        )
        if dir_path:
            self.pdf_dir_edit.setText(dir_path)
            self.saveSettings()

    def load_csv_data(self, file_path):
        try:
            self.log_text.append_log(
                LogMessage(f"Carregando arquivo: {file_path}", LogMessage.INFO)
            )

            # Detect file type by extension
            file_ext = os.path.splitext(file_path)[1].lower()

            if file_ext in [".xlsx", ".xls"]:
                # Excel file
                self.log_text.append_log(
                    LogMessage(
                        "Detectado arquivo Excel, processando...", LogMessage.INFO
                    )
                )
                try:
                    data = pd.read_excel(file_path, dtype=str)
                    self.log_text.append_log(
                        LogMessage(
                            "Arquivo Excel carregado com sucesso", LogMessage.SUCCESS
                        )
                    )
                except Exception as excel_err:
                    self.log_text.append_log(
                        LogMessage(
                            f"Erro ao processar Excel: {str(excel_err)}",
                            LogMessage.ERROR,
                        )
                    )
                    QMessageBox.critical(
                        self,
                        "Erro",
                        f"Erro ao processar arquivo Excel:\n{str(excel_err)}",
                    )
                    return
            else:
                # CSV file - try to detect encoding and delimiter
                try:
                    # First check for comment line
                    with open(file_path, "r", errors="ignore") as f:
                        first_line = f.readline().strip()
                        skip_rows = 1 if first_line.startswith("//") else 0

                    # Try to detect delimiter
                    with open(file_path, "r", errors="ignore") as f:
                        sample = f.read(1024)
                        if sample.count(";") > sample.count(","):
                            delimiter = ";"
                        else:
                            delimiter = ","

                    # Try different encodings
                    encodings = ["utf-8", "latin1", "cp1252", "utf-8-sig"]
                    for encoding in encodings:
                        try:
                            data = pd.read_csv(
                                file_path,
                                dtype=str,
                                skiprows=skip_rows,
                                sep=delimiter,
                                encoding=encoding,
                            )
                            self.log_text.append_log(
                                LogMessage(
                                    f"CSV carregado com encoding {encoding} e delimitador '{delimiter}'",
                                    LogMessage.SUCCESS,
                                )
                            )
                            break
                        except UnicodeDecodeError:
                            continue
                        except Exception as e:
                            raise e
                    else:
                        # If all encodings fail
                        raise Exception(
                            "Não foi possível determinar a codificação do arquivo CSV"
                        )
                except Exception as csv_err:
                    self.log_text.append_log(
                        LogMessage(
                            f"Erro ao processar CSV: {str(csv_err)}", LogMessage.ERROR
                        )
                    )
                    QMessageBox.critical(
                        self, "Erro", f"Erro ao processar arquivo CSV:\n{str(csv_err)}"
                    )
                    return

            # Clean column names by stripping whitespace
            data.columns = data.columns.str.strip()

            # Map column names
            column_mapping = {
                "CPF/CNPJ": "CPF_CNPJ",
                "SERVIÇO": "SERVICO",
                "REFERENCIA": "REFERENCIA",
                "VENCIMENTO": "VENCIMENTO",
                "VALOR": "VALOR",
                "NOTA FISCAL": "NF",  # Changed from "nf" to "NOTA FISCAL"
                "INFORMAÇÕES ADICIONAIS": "INFO_ADICIONAIS",
            }

            # Check if required columns exist
            required_columns = [
                "CPF/CNPJ",
                "SERVIÇO",
                "REFERENCIA",
                "VENCIMENTO",
                "VALOR",
            ]
            missing_columns = []

            for req_col in required_columns:
                # Check if column exists (case insensitive comparison)
                if not any(
                    col.strip().upper() == req_col.upper() for col in data.columns
                ):
                    # Try alternate names
                    alternates = {
                        "CPF/CNPJ": ["CPF_CNPJ", "CNPJ", "CPF", "DOCUMENTO"],
                        "SERVIÇO": ["SERVICO", "CODIGO", "SERV", "COD_SERVICO"],
                        "REFERENCIA": ["REF", "PERIODO", "COMPETENCIA"],
                        "VENCIMENTO": ["DATA_VENC", "DT_VENCIMENTO", "VENC"],
                        "VALOR": ["VLR", "TOTAL", "MONTANTE"],
                    }

                    found = False
                    if req_col in alternates:
                        for alt in alternates[req_col]:
                            if any(
                                col.strip().upper() == alt.upper()
                                for col in data.columns
                            ):
                                # Rename in the dataframe
                                for i, col in enumerate(data.columns):
                                    if col.strip().upper() == alt.upper():
                                        # Rename in the dataframe
                                        data.rename(
                                            columns={col: req_col}, inplace=True
                                        )
                                        found = True
                                        break
                                if found:
                                    break

                    if not found:
                        missing_columns.append(req_col)

            if missing_columns:
                error_msg = f"Colunas obrigatórias não encontradas: {', '.join(missing_columns)}"
                self.log_text.append_log(LogMessage(error_msg, LogMessage.ERROR))
                QMessageBox.critical(self, "Erro", error_msg)
                return

            # Rename columns
            data = data.rename(columns=column_mapping)

            # Clean data
            data["VALOR"] = data["VALOR"].str.strip().str.replace(",", ".")
            data["VALOR"] = data["VALOR"].str.replace(" ", "")

            # Ensure NF column exists
            if "NF" not in data.columns:
                data["NF"] = ""

            # Ensure INFO_ADICIONAIS column exists
            if "INFO_ADICIONAIS" not in data.columns:
                data["INFO_ADICIONAIS"] = ""

            # Combine fields
            data["INFO_COMBINADA"] = (
                "NF: "
                + data["NF"].fillna("")
                + " - "
                + data["INFO_ADICIONAIS"].fillna("")
            )

            # Update table
            self.table_model.update_data(data)
            self.data = data

            # Enable start button
            self.start_btn.setEnabled(True)

            # Update status
            self.status_label.setText(
                f"Status: Arquivo carregado com {len(data)} registros"
            )
            self.statusBar.showMessage(f"Carregado: {len(data)} registros")
            self.log_text.append_log(
                LogMessage(
                    f"Arquivo carregado com sucesso: {len(data)} registros",
                    LogMessage.SUCCESS,
                )
            )

            # Log the first few records as a preview
            if len(data) > 0:
                self.log_text.append_log(
                    LogMessage("Prévia dos dados:", LogMessage.INFO)
                )
                for i, row in data.head(3).iterrows():
                    self.log_text.append_log(
                        LogMessage(
                            f"  Item {i+1}: CPF/CNPJ: {row['CPF_CNPJ']}, "
                            f"Ref: {row['REFERENCIA']}, Valor: {row['VALOR']}",
                            LogMessage.INFO,
                        )
                    )
                if len(data) > 3:
                    self.log_text.append_log(
                        LogMessage(
                            f"  ... e mais {len(data) - 3} registros", LogMessage.INFO
                        )
                    )

        except Exception as e:
            self.log_text.append_log(
                LogMessage(f"Erro ao carregar o arquivo: {str(e)}", LogMessage.ERROR)
            )
            QMessageBox.critical(self, "Erro", f"Erro ao carregar o arquivo:\n{str(e)}")

    def start_processing(self):
        if not hasattr(self, "data") or self.data.empty:
            self.log_text.append_log(
                LogMessage("⚠️ Nenhum dado para processar.", LogMessage.WARNING)
            )
            QMessageBox.warning(self, "Aviso", "Nenhum dado para processar.")
            return

        # Update UI state
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.select_file_btn.setEnabled(False)

        # Initialize progress bar
        total_rows = len(self.data)
        self.progress_bar.setRange(0, total_rows)
        self.progress_bar.setValue(0)

        # Reset counters
        self.success_count.setText("Sucessos: 0")
        self.failure_count.setText("Falhas: 0")

        # Update status
        self.status_label.setText("Status: Iniciando processamento...")

        # Clear log
        self.log_text.clear()
        self.log_text.append_log(
            LogMessage("🚀 Iniciando processamento de DUAs...", LogMessage.INFO)
        )
        self.log_text.append_log(
            LogMessage(
                f"📋 Total de registros a processar: {total_rows}", LogMessage.INFO
            )
        )

        # Get PDF directory
        pdf_dir = self.pdf_dir_edit.text()
        os.makedirs(pdf_dir, exist_ok=True)
        self.log_text.append_log(
            LogMessage(f"📁 PDFs serão salvos em: {pdf_dir}", LogMessage.INFO)
        )

        # Start worker thread
        self.worker = WorkerThread(self.data, pdf_dir)
        self.worker.progress_signal.connect(self.update_progress)
        self.worker.finished_signal.connect(self.process_finished)
        self.worker.log_signal.connect(self.update_log)
        self.worker.status_signal.connect(self.update_status)
        self.worker.captcha_signal.connect(
            self.show_captcha_dialog
        )  # Conectar o sinal de captcha
        self.worker.start()

    def stop_processing(self):
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self,
                "Confirmar",
                "Deseja realmente interromper o processamento?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )

            if reply == QMessageBox.StandardButton.Yes:
                self.log_text.append_log(
                    LogMessage("⚠️ Interrompendo processamento...", LogMessage.WARNING)
                )
                self.status_label.setText("Status: Interrompendo processamento...")

                # Improved stop operation
                try:
                    self.worker.stop()
                    # Signal UI that we're stopping
                    self.stop_btn.setEnabled(False)
                    self.stop_btn.setText("Parando...")

                    # Give visual feedback that stop was requested
                    self.statusBar.showMessage(
                        "Aguardando término das operações em andamento..."
                    )

                    # Optional: Force quit after timeout (uncomment if needed)
                    # QTimer.singleShot(10000, self.force_stop)
                except Exception as e:
                    self.log_text.append_log(
                        LogMessage(f"Erro ao tentar parar: {str(e)}", LogMessage.ERROR)
                    )

    def force_stop(self):
        """Emergency force stop if regular stop doesn't work"""
        if self.worker and self.worker.isRunning():
            try:
                # Try to terminate the thread (more forceful)
                self.worker.terminate()
                self.log_text.append_log(
                    LogMessage(
                        "Processamento terminado forçadamente", LogMessage.WARNING
                    )
                )

                # Update UI
                self.start_btn.setEnabled(True)
                self.stop_btn.setEnabled(False)
                self.stop_btn.setText("Parar")
                self.select_file_btn.setEnabled(True)
                self.statusBar.showMessage("Processamento interrompido forçadamente")

                # Close browser if it's still open
                from get_dua import close_browser

                close_browser()
            except Exception as e:
                self.log_text.append_log(
                    LogMessage(f"Erro ao forçar parada: {str(e)}", LogMessage.ERROR)
                )

    def update_progress(self, current, total):
        self.progress_bar.setValue(current)
        percentage = int(current / total * 100)
        self.statusBar.showMessage(f"Processando... {percentage}%")

    def update_log(self, message):
        self.log_text.append_log(message)
        # Update counters based on message type
        if message.level == LogMessage.SUCCESS and "✅" in message.text:
            current = int(self.success_count.text().split(":")[1])
            self.success_count.setText(f"Sucessos: {current + 1}")
        elif message.level == LogMessage.ERROR and "❌" in message.text:
            current = int(self.failure_count.text().split(":")[1])
            self.failure_count.setText(f"Falhas: {current + 1}")

    def update_status(self, message, level):
        # Update the status label with color based on level
        self.status_label.setText(f"Status: {message}")
        if level == LogMessage.ERROR:
            self.status_label.setStyleSheet("font-weight: bold; color: red;")
        elif level == LogMessage.WARNING:
            self.status_label.setStyleSheet("font-weight: bold; color: orange;")
        elif level == LogMessage.SUCCESS:
            self.status_label.setStyleSheet("font-weight: bold; color: green;")
        else:
            self.status_label.setStyleSheet("font-weight: bold;")

    def show_captcha_dialog(self):
        """Exibe o diálogo de resolução manual de CAPTCHA"""
        self.log_text.append_log(
            LogMessage(
                "⚠️ O CAPTCHA não pôde ser resolvido automaticamente!",
                LogMessage.WARNING,
            )
        )
        self.log_text.append_log(
            LogMessage(
                "🔐 Por favor, resolva o CAPTCHA manualmente no navegador.",
                LogMessage.WARNING,
            )
        )

        # Importar apenas quando necessário
        from get_dua import captcha_solved_signal

        # Criar e exibir o diálogo
        captcha_dialog = CaptchaDialog(
            self, captcha_solved_callback=captcha_solved_signal
        )
        result = captcha_dialog.exec()

        if result == QDialog.DialogCode.Accepted:
            self.log_text.append_log(
                LogMessage(
                    "✅ CAPTCHA resolvido manualmente pelo usuário.", LogMessage.SUCCESS
                )
            )
        else:
            self.log_text.append_log(
                LogMessage(
                    "❌ Resolução de CAPTCHA cancelada pelo usuário.", LogMessage.ERROR
                )
            )
            self.stop_processing()

    def process_finished(self, success):
        self.start_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
        self.select_file_btn.setEnabled(True)

        if success:
            self.statusBar.showMessage("Processamento concluído com sucesso!")
            self.status_label.setStyleSheet("font-weight: bold; color: green;")
            self.status_label.setText("Status: Processamento concluído com sucesso!")
        else:
            self.statusBar.showMessage("Processamento concluído com falhas.")
            self.status_label.setStyleSheet("font-weight: bold; color: orange;")
            self.status_label.setText("Status: Processamento concluído com falhas.")

        self.log_text.append_log(
            LogMessage("Processamento concluído!", LogMessage.SUCCESS)
        )

    def closeEvent(self, event):
        self.saveSettings()

        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self,
                "Confirmar",
                "O processamento está em andamento. Deseja realmente sair?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )

            if reply == QMessageBox.StandardButton.Yes:
                self.worker.stop()
                self.worker.wait()
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # Modern look across platforms
    window = DUAAutomationUI()
    window.show()
    sys.exit(app.exec())
