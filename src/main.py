from dataclasses import dataclass
import os
from pathlib import Path
import re
import sys
import threading
from typing import Any, List, Optional

from PySide6.QtCore import QObject, Qt, Signal
from PySide6.QtGui import QColor, QIcon
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHeaderView,
    QHBoxLayout,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)
from qfluentwidgets import (
    CaptionLabel,
    CheckBox,
    ComboBox,
    FluentIcon as FIF,
    InfoBar,
    InfoBarPosition,
    MSFluentWindow,
    ProgressBar,
    PushButton,
    PrimaryPushButton,
    SubtitleLabel,
    TableWidget,
    Theme,
    setTheme,
    setThemeColor,
)

if sys.platform == "win32":
    import ctypes
    import winreg
    import pythoncom
    import win32com.client

    PLATFORM_WINDOWS = True
else:
    PLATFORM_WINDOWS = False


@dataclass(frozen=True)
class AttachmentItem:
    filename: str
    date: str
    subject: str
    entry_id: str
    att_index: int


def get_windows_accent_color() -> QColor:
    if not PLATFORM_WINDOWS:
        return QColor(0, 120, 212)
    try:
        registry = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
        key = winreg.OpenKey(registry, r"Software\Microsoft\Windows\DWM")
        value, _ = winreg.QueryValueEx(key, "AccentColor")
        r = value & 0xFF
        g = (value >> 8) & 0xFF
        b = (value >> 16) & 0xFF
        return QColor(r, g, b)
    except Exception:
        return QColor(0, 120, 212)


def init_windows_app(app_id: str) -> Path:
    if PLATFORM_WINDOWS:
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
        except Exception:
            pass

    base_path = (
        Path(sys._MEIPASS) if getattr(sys, "frozen", False) else Path(__file__).parent
    )
    icon_path = base_path / "assets/mail.ico"
    return icon_path if icon_path.exists() else Path()


class SyncSignals(QObject):
    email_found = Signal(dict)
    finished = Signal()
    progress = Signal(int)


class SyncWorker:
    def __init__(self, account_name: str, folder_name: str) -> None:
        self.account_name = account_name
        self.folder_name = folder_name
        self.signals = SyncSignals()

    def run(self) -> None:
        if not PLATFORM_WINDOWS:
            self.signals.finished.emit()
            return

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
                        for att in msg.Attachments:
                            data = {
                                "filename": att.FileName,
                                "date": str(msg.ReceivedTime)[:16],
                                "subject": str(msg.Subject),
                                "entry_id": msg.EntryID,
                                "att_index": att.Index,
                            }
                            self.signals.email_found.emit(data)
                    self.signals.progress.emit(int((i / total) * 100))
        except Exception as e:
            print(f"Sync Error: {e}")
        finally:
            self.signals.finished.emit()
            pythoncom.CoUninitialize()


class OutlookDownloader(QWidget):
    def __init__(self, parent: Optional[QWidget] = None) -> None:
        super().__init__(parent=parent)
        self.setObjectName("OutlookDownloader")
        self.current_items: List[AttachmentItem] = []
        self.is_dark = True
        self.namespace: Any = None

        if PLATFORM_WINDOWS:
            self.outlook_app = win32com.client.Dispatch("Outlook.Application")
            self.namespace = self.outlook_app.GetNamespace("MAPI")

        self.setup_ui()
        self.load_accounts()

    def setup_ui(self) -> None:
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(30, 30, 30, 30)
        self.main_layout.setSpacing(20)

        header_layout = QHBoxLayout()
        title_container = QVBoxLayout()
        self.title_label = SubtitleLabel("Anexotlook")
        self.subtitle_label = CaptionLabel("Sistema de extração de anexos")
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
        acc_box.addWidget(CaptionLabel("Contas"))
        self.combo_accounts = ComboBox()
        self.combo_accounts.currentIndexChanged.connect(self.on_account_changed)
        acc_box.addWidget(self.combo_accounts)

        fld_box = QVBoxLayout()
        fld_box.addWidget(CaptionLabel("Pasta"))
        self.combo_folders = ComboBox()
        fld_box.addWidget(self.combo_folders)

        self.btn_sync = PrimaryPushButton(FIF.SYNC, "Sincronizar")
        self.btn_sync.setFixedWidth(150)
        self.btn_sync.clicked.connect(self.start_sync)

        filter_layout.addLayout(acc_box, 2)
        filter_layout.addLayout(fld_box, 2)
        filter_layout.addWidget(self.btn_sync, 0, Qt.AlignmentFlag.AlignBottom)
        self.main_layout.addLayout(filter_layout)

        self.progress_bar = ProgressBar()
        self.progress_bar.setHidden(True)
        self.main_layout.addWidget(self.progress_bar)

        self.main_layout.addWidget(CaptionLabel("Encontrado"))
        self.table = TableWidget()
        self.table.setColumnCount(4)
        self.table.setHorizontalHeaderLabels(
            ["Sel.", "Nome do Arquivo", "Data do E-mail", "Assunto"]
        )
        self.table.horizontalHeader().setSectionResizeMode(
            QHeaderView.ResizeMode.Stretch
        )
        self.table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.ResizeToContents
        )
        self.main_layout.addWidget(self.table)

        footer = QHBoxLayout()
        self.lbl_status = CaptionLabel("Status: Aguardando")
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

    def toggle_theme(self) -> None:
        self.is_dark = not self.is_dark
        setTheme(Theme.DARK if self.is_dark else Theme.LIGHT)

    def load_accounts(self) -> None:
        if PLATFORM_WINDOWS and self.namespace:
            try:
                accounts = [acc.DisplayName for acc in self.namespace.Accounts]
                self.combo_accounts.addItems(accounts)
            except Exception:
                pass

    def on_account_changed(self) -> None:
        acc_name = self.combo_accounts.currentText()
        self.combo_folders.clear()
        if PLATFORM_WINDOWS and self.namespace:
            for acc in self.namespace.Folders:
                if acc.Name == acc_name:
                    for folder in acc.Folders:
                        self.combo_folders.addItem(folder.Name)
                    break

    def start_sync(self) -> None:
        acc = self.combo_accounts.currentText()
        fld = self.combo_folders.currentText()
        if not acc or not fld:
            return

        self.table.setRowCount(0)
        self.current_items.clear()
        self.btn_sync.setEnabled(False)
        self.progress_bar.setHidden(False)
        self.lbl_status.setText("Status: Analisando...")

        self.worker = SyncWorker(acc, fld)
        self.worker.signals.email_found.connect(self.add_file_to_table)
        self.worker.signals.progress.connect(self.progress_bar.setValue)
        self.worker.signals.finished.connect(self.on_sync_finished)
        threading.Thread(target=self.worker.run, daemon=True).start()

    def add_file_to_table(self, data: dict) -> None:
        item = AttachmentItem(**data)
        self.current_items.append(item)

        row = self.table.rowCount()
        self.table.insertRow(row)

        chk = CheckBox()
        h_layout = QHBoxLayout()
        h_layout.addWidget(chk)
        h_layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        h_layout.setContentsMargins(0, 0, 0, 0)
        container = QWidget()
        container.setLayout(h_layout)

        self.table.setCellWidget(row, 0, container)
        self.table.setItem(row, 1, QTableWidgetItem(item.filename))
        self.table.setItem(row, 2, QTableWidgetItem(item.date))
        self.table.setItem(row, 3, QTableWidgetItem(item.subject))

    def on_sync_finished(self) -> None:
        self.btn_sync.setEnabled(True)
        self.progress_bar.setHidden(True)
        self.lbl_status.setText(
            f"Status: {len(self.current_items)} Encontrado"
        )
        InfoBar.success(
            "Concluído",
            "Busca finalizada.",
            position=InfoBarPosition.TOP,
            parent=self,
        )

    def toggle_select_all(self) -> None:
        for r in range(self.table.rowCount()):
            container = self.table.cellWidget(r, 0)
            if container and container.layout():
                chk = container.layout().itemAt(0).widget()
                if isinstance(chk, CheckBox):
                    chk.setChecked(True)

    def download_selected(self) -> None:
        selected_indices = []
        for r in range(self.table.rowCount()):
            container = self.table.cellWidget(r, 0)
            if container and container.layout():
                chk = container.layout().itemAt(0).widget()
                if isinstance(chk, CheckBox) and chk.isChecked():
                    selected_indices.append(r)

        if not selected_indices:
            InfoBar.warning("Aviso", "Selecione pelo menos um arquivo.", parent=self)
            return

        destination_dir = QFileDialog.getExistingDirectory(
            self, "Selecionar Pasta de Destino"
        )
        if not destination_dir:
            return

        target_path = Path(destination_dir)

        try:
            count = 0
            for idx in selected_indices:
                item = self.current_items[idx]
                msg = self.namespace.GetItemFromID(item.entry_id)
                att = msg.Attachments.Item(item.att_index)

                clean_name = re.sub(r'[\\/*?:"<>|]', "", att.FileName)
                save_path = target_path / clean_name

                if save_path.exists():
                    name_part = save_path.stem
                    ext = save_path.suffix
                    save_path = target_path / f"{name_part}_{idx}{ext}"

                att.SaveAsFile(str(save_path))
                count += 1

            InfoBar.success(
                "Sucesso", f"{count} arquivos salvos com sucesso!", parent=self
            )
        except Exception as e:
            InfoBar.error("Erro no Download", str(e), parent=self)


class MainWindow(MSFluentWindow):
    def __init__(self, icon_path: Path) -> None:
        super().__init__()
        accent_color = get_windows_accent_color()
        setThemeColor(accent_color)
        setTheme(Theme.AUTO)

        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))

        self.downloader_widget = OutlookDownloader(self)
        self.addSubInterface(self.downloader_widget, FIF.MAIL, "Arquivos")
        self.setWindowTitle("Anexotlook")
        self.resize(1100, 800)


def main() -> None:
    app = QApplication(sys.argv)
    icon_p = init_windows_app("gui23x.anexotlook.v2")
    window = MainWindow(icon_p)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
