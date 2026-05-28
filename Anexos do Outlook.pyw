import sys
import os
import re
import threading
import ctypes
import winreg  # Para buscar a cor do Windows
from PySide6.QtWidgets import (
    QApplication,
    QVBoxLayout,
    QHBoxLayout,
    QHeaderView,
    QFileDialog,
    QTableWidgetItem,
    QWidget,
    QFrame,
)
from PySide6.QtCore import Qt, Signal, QObject
from PySide6.QtGui import QIcon, QColor

# Importando Componentes Fluent Design
from qfluentwidgets import (
    PushButton,
    PrimaryPushButton,
    ComboBox,
    TableWidget,
    ProgressBar,
    CheckBox,
    SubtitleLabel,
    CaptionLabel,
    setTheme,
    Theme,
    setThemeColor,
    InfoBar,
    InfoBarPosition,
    MSFluentWindow,
    FluentIcon as FIF,
)

try:
    import win32com.client
    import pythoncom

    PLATFORM_WINDOWS = True
except ImportError:
    PLATFORM_WINDOWS = False


# --- FUNÇÃO PARA PEGAR A COR DE DESTAQUE DO WINDOWS ---
def get_windows_accent_color():
    """Busca a cor de destaque (Accent Color) configurada no Windows 11"""
    try:
        registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        key = winreg.OpenKey(registry, r"Software\Microsoft\Windows\DWM")
        value, _ = winreg.QueryValueEx(key, "AccentColor")
        # O valor vem em formato ABGR (hex), precisamos converter para RGB
        # O Windows armazena como um inteiro (ex: 4289390164)
        a = (value >> 24) & 0xFF
        b = (value >> 16) & 0xFF
        g = (value >> 8) & 0xFF
        r = value & 0xFF
        return QColor(r, g, b)
    except Exception:
        return QColor(0, 120, 212)  # Azul padrão Windows caso falhe


# --- CONFIGURAÇÃO DE SISTEMA (ÍCONE E BARRA DE TAREFAS) ---
def init_windows_app(app_id):
    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass

    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))

    icon_path = os.path.join(base_path, "mail.ico")
    return icon_path if os.path.exists(icon_path) else ""


class SyncSignals(QObject):
    email_found = Signal(dict)
    finished = Signal()
    progress = Signal(int)


class SyncWorker:
    def __init__(self, account_name, folder_name):
        self.account_name = account_name
        self.folder_name = folder_name
        self.signals = SyncSignals()

    def run(self):
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            target_folder = None
            for acc in namespace.Folders:
                if acc.Name == self.account_name:
                    for folder in acc.Folders:
                        if folder.Name == self.folder_name:
                            target_folder = folder
                            break
                    break

            if target_folder:
                messages = target_folder.Items
                messages.Sort("[ReceivedTime]", True)
                total = min(len(messages), 250)
                for i in range(1, total + 1):
                    msg = messages.Item(i)
                    if hasattr(msg, "Attachments") and msg.Attachments.Count > 0:
                        data = {
                            "subject": str(msg.Subject),
                            "date": str(msg.ReceivedTime)[:16],
                            "attachments": [att.FileName for att in msg.Attachments],
                            "entry_id": msg.EntryID,
                        }
                        self.signals.email_found.emit(data)
                    self.signals.progress.emit(int((i / total) * 100))
        except Exception as e:
            print(f"Erro Sync: {e}")
        self.signals.finished.emit()
        pythoncom.CoUninitialize()


class OutlookDownloader(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent=parent)
        self.setObjectName("OutlookDownloader")
        self.current_emails = []
        self.is_dark = True

        if PLATFORM_WINDOWS:
            self.outlook_app = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook_app.GetNamespace("MAPI")

        self.setup_ui()
        self.load_accounts()

    def setup_ui(self):
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(30, 30, 30, 30)
        self.main_layout.setSpacing(20)

        header_layout = QHBoxLayout()
        title_container = QVBoxLayout()
        self.title_label = SubtitleLabel("Outlook Downloader")
        self.subtitle_label = CaptionLabel("SISTEMA DE SINCRONIZAÇÃO DE ANEXOS")
        title_container.addWidget(self.title_label)
        title_container.addWidget(self.subtitle_label)

        self.btn_theme = PushButton(FIF.PALETTE, "Mudar Tema")
        self.btn_theme.clicked.connect(self.toggle_theme)

        header_layout.addLayout(title_container)
        header_layout.addStretch()
        header_layout.addWidget(self.btn_theme)
        self.main_layout.addLayout(header_layout)

        filter_layout = QHBoxLayout()
        acc_box = QVBoxLayout()
        acc_box.addWidget(CaptionLabel("CONTA"))
        self.combo_accounts = ComboBox()
        self.combo_accounts.currentIndexChanged.connect(self.on_account_changed)
        acc_box.addWidget(self.combo_accounts)

        fld_box = QVBoxLayout()
        fld_box.addWidget(CaptionLabel("PASTA"))
        self.combo_folders = ComboBox()
        fld_box.addWidget(self.combo_folders)

        self.btn_sync = PrimaryPushButton(FIF.SYNC, "Sincronizar")
        self.btn_sync.setFixedWidth(150)
        self.btn_sync.clicked.connect(self.start_sync)

        filter_layout.addLayout(acc_box, 2)
        filter_layout.addLayout(fld_box, 2)
        filter_layout.addWidget(self.btn_sync, 0, Qt.AlignBottom)
        self.main_layout.addLayout(filter_layout)

        self.progress_bar = ProgressBar()
        self.progress_bar.setHidden(True)
        self.main_layout.addWidget(self.progress_bar)

        self.main_layout.addWidget(CaptionLabel("MENSAGENS DETECTADAS"))
        self.table = TableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(["Sel.", "Assunto", "Data", "Anexos"])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeToContents
        )
        self.main_layout.addWidget(self.table)

        footer = QHBoxLayout()
        self.lbl_status = CaptionLabel("STATUS: AGUARDANDO")
        self.btn_select_all = PushButton("Selecionar Todos")
        self.btn_select_all.clicked.connect(self.toggle_select_all)
        self.btn_download = PrimaryPushButton(FIF.DOWNLOAD, "Baixar Selecionados")
        self.btn_download.setFixedWidth(200)
        self.btn_download.clicked.connect(self.download_selected)

        footer.addWidget(self.lbl_status)
        footer.addStretch()
        footer.addWidget(self.btn_select_all)
        footer.addWidget(self.btn_download)
        self.main_layout.addLayout(footer)

    def toggle_theme(self):
        self.is_dark = not self.is_dark
        setTheme(Theme.DARK if self.is_dark else Theme.LIGHT)

    def load_accounts(self):
        if PLATFORM_WINDOWS:
            try:
                accounts = [acc.DisplayName for acc in self.namespace.Accounts]
                self.combo_accounts.addItems(accounts)
            except:
                pass

    def on_account_changed(self):
        acc_name = self.combo_accounts.currentText()
        self.combo_folders.clear()
        if PLATFORM_WINDOWS:
            for acc in self.namespace.Folders:
                if acc.Name == acc_name:
                    for folder in acc.Folders:
                        self.combo_folders.addItem(folder.Name)
                    break

    def start_sync(self):
        acc = self.combo_accounts.currentText()
        fld = self.combo_folders.currentText()
        if not acc or not fld:
            return
        self.table.setRowCount(0)
        self.current_emails = []
        self.btn_sync.setEnabled(False)
        self.progress_bar.setHidden(False)
        self.lbl_status.setText("STATUS: SINCRONIZANDO...")
        self.worker = SyncWorker(acc, fld)
        self.worker.signals.email_found.connect(self.add_email_to_table)
        self.worker.signals.progress.connect(self.progress_bar.setValue)
        self.worker.signals.finished.connect(self.on_sync_finished)
        threading.Thread(target=self.worker.run, daemon=True).start()

    def add_email_to_table(self, data):
        self.current_emails.append(data)
        row = self.table.rowCount()
        self.table.insertRow(row)
        chk = CheckBox()
        h_layout = QHBoxLayout()
        h_layout.addWidget(chk)
        h_layout.setAlignment(Qt.AlignCenter)
        h_layout.setContentsMargins(0, 0, 0, 0)
        container = QWidget()
        container.setLayout(h_layout)
        self.table.setCellWidget(row, 0, container)
        self.table.setItem(row, 1, QTableWidgetItem(data["subject"]))
        self.table.setItem(row, 2, QTableWidgetItem(data["date"]))
        self.table.setItem(row, 3, QTableWidgetItem(f"{len(data['attachments'])} arq."))

    def on_sync_finished(self):
        self.btn_sync.setEnabled(True)
        self.progress_bar.setHidden(True)
        self.lbl_status.setText(
            f"STATUS: {len(self.current_emails)} EMAILS ENCONTRADOS"
        )
        InfoBar.success(
            title="Concluído",
            content="Sincronização finalizada.",
            position=InfoBarPosition.TOP,
            parent=self,
        )

    def toggle_select_all(self):
        for r in range(self.table.rowCount()):
            container = self.table.cellWidget(r, 0)
            chk = container.layout().itemAt(0).widget()
            chk.setChecked(True)

    def download_selected(self):
        selected_data = []
        for r in range(self.table.rowCount()):
            container = self.table.cellWidget(r, 0)
            chk = container.layout().itemAt(0).widget()
            if chk.isChecked():
                selected_data.append(self.current_emails[r])
        if not selected_data:
            return InfoBar.warning(
                title="Aviso", content="Selecione um item.", parent=self
            )
        path = QFileDialog.getExistingDirectory(self, "Salvar em...")
        if not path:
            return
        try:
            count = 0
            for data in selected_data:
                msg = self.namespace.GetItemFromID(data["entry_id"])
                for att in msg.Attachments:
                    name = re.sub(r'[\\/*?:"<>|]', "", att.FileName)
                    att.SaveAsFile(os.path.join(path, name))
                    count += 1
            InfoBar.success(
                title="Sucesso", content=f"{count} arquivos salvos.", parent=self
            )
        except Exception as e:
            InfoBar.error(title="Erro", content=str(e), parent=self)


class MainWindow(MSFluentWindow):
    def __init__(self, icon_path):
        super().__init__()

        # BUSCAR E APLICAR COR DO SISTEMA OPERACIONAL
        accent_color = get_windows_accent_color()
        setThemeColor(accent_color)

        setTheme(Theme.AUTO)

        if icon_path:
            self.setWindowIcon(QIcon(icon_path))

        self.downloader_widget = OutlookDownloader(self)
        self.addSubInterface(self.downloader_widget, FIF.MAIL, "Downloads")

        self.setWindowTitle("Outlook Downloader Pro")
        self.resize(1100, 800)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    icon_p = init_windows_app("meu.outlook.downloader.v1")
    window = MainWindow(icon_p)
    window.show()
    sys.exit(app.exec())
