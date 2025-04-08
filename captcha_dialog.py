from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, 
    QPushButton, QGroupBox, QMessageBox
)
from PyQt6.QtCore import Qt, QTimer
from PyQt6.QtGui import QFont, QIcon, QPixmap

class CaptchaDialog(QDialog):
    """Diálogo para resolução manual de CAPTCHA"""
    
    def __init__(self, parent=None, captcha_solved_callback=None):
        super().__init__(parent)
        self.captcha_solved_callback = captcha_solved_callback
        self.setWindowFlags(Qt.WindowType.Dialog | Qt.WindowType.WindowStaysOnTopHint)
        self.setWindowModality(Qt.WindowModality.ApplicationModal)
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle("Resolução Manual de CAPTCHA")
        self.setMinimumWidth(500)
        self.setStyleSheet("""
            QDialog {
                background-color: #f0f2f5;
            }
            QGroupBox {
                border: 1px solid #d0d0d0;
                border-radius: 5px;
                margin-top: 1em;
                padding-top: 10px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                left: 10px;
                padding: 0 3px 0 3px;
                font-weight: bold;
            }
            QPushButton {
                background-color: #4a86e8;
                color: white;
                border: none;
                padding: 8px 16px;
                border-radius: 4px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #5a96f8;
            }
            QPushButton:pressed {
                background-color: #3a76d8;
            }
        """)
        
        layout = QVBoxLayout()
        
        # Título com ícone de atenção
        title_layout = QHBoxLayout()
        
        attention_icon = QLabel()
        try:
            # Tentar carregar um ícone de alerta
            icon_path = "resources/warning_icon.png"
            pixmap = QPixmap(icon_path)
            if not pixmap.isNull():
                attention_icon.setPixmap(pixmap.scaled(32, 32, Qt.AspectRatioMode.KeepAspectRatio))
                title_layout.addWidget(attention_icon)
        except:
            pass  # Ignorar erro se o ícone não for encontrado
            
        title_label = QLabel("Ação necessária: Resolução de CAPTCHA")
        title_label.setFont(QFont("Arial", 14, QFont.Weight.Bold))
        title_label.setStyleSheet("color: #e63946;")
        title_layout.addWidget(title_label)
        title_layout.addStretch()
        
        layout.addLayout(title_layout)
        
        # Instruções
        instructions_group = QGroupBox("Instruções")
        instructions_layout = QVBoxLayout(instructions_group)
        
        instructions_text = QLabel(
            "<p>O sistema não conseguiu resolver o CAPTCHA automaticamente.</p>"
            "<p><b>Por favor, siga os passos abaixo:</b></p>"
            "<ol>"
            "<li>Localize a janela do navegador Chrome aberta pelo sistema</li>"
            "<li>Observe o formulário de DUA com o CAPTCHA sendo exibido</li>"
            "<li>Resolva o CAPTCHA visualmente, clicando nas imagens solicitadas</li>"
            "<li>Após resolver o CAPTCHA, retorne a este diálogo</li>"
            "<li>Clique no botão 'CAPTCHA Resolvido' abaixo</li>"
            "</ol>"
            "<p style='color: #e63946;'><b>IMPORTANTE:</b> Não feche esta janela até resolver o CAPTCHA.</p>"
        )
        instructions_text.setWordWrap(True)
        instructions_layout.addWidget(instructions_text)
        
        layout.addWidget(instructions_group)
        
        # Timer de espera
        self.wait_label = QLabel("Aguardando resolução do CAPTCHA...")
        self.wait_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.wait_label.setStyleSheet("font-style: italic; color: #666;")
        layout.addWidget(self.wait_label)
        
        # Botões
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.solved_button = QPushButton("CAPTCHA Resolvido")
        self.solved_button.setMinimumWidth(150)
        self.solved_button.clicked.connect(self.on_captcha_solved)
        button_layout.addWidget(self.solved_button)
        
        self.cancel_button = QPushButton("Cancelar")
        self.cancel_button.setStyleSheet("background-color: #d32f2f;")
        self.cancel_button.clicked.connect(self.on_cancel)
        button_layout.addWidget(self.cancel_button)
        
        layout.addLayout(button_layout)
        
        self.setLayout(layout)
        
    def on_captcha_solved(self):
        """Callback quando o usuário clica em 'CAPTCHA Resolvido'"""
        if self.captcha_solved_callback:
            self.captcha_solved_callback()
        self.accept()
        
    def on_cancel(self):
        """Callback quando o usuário cancela a operação"""
        reply = QMessageBox.question(
            self, 
            "Confirmar cancelamento", 
            "Cancelar a resolução do CAPTCHA interromperá o processamento atual. Deseja continuar?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )
        
        if reply == QMessageBox.StandardButton.Yes:
            self.reject()
        
    def showEvent(self, event):
        """Override para iniciar animação quando o diálogo aparecer"""
        super().showEvent(event)
        
        # Iniciar timer para animação de aguardando
        self.dot_count = 0
        self.wait_timer = QTimer(self)
        self.wait_timer.timeout.connect(self.update_waiting_text)
        self.wait_timer.start(500)  # Atualiza a cada meio segundo
        
    def update_waiting_text(self):
        """Atualiza o texto de aguardando com animação de pontos"""
        self.dot_count = (self.dot_count % 3) + 1
        dots = "." * self.dot_count
        self.wait_label.setText(f"Aguardando resolução do CAPTCHA{dots}")
        
    def closeEvent(self, event):
        """Override para limpar o timer ao fechar"""
        if hasattr(self, 'wait_timer'):
            self.wait_timer.stop()
        super().closeEvent(event)
