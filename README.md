# Outlook Attachment Downloader

<img src="https://i.pinimg.com/originals/25/0f/01/250f01e48a6d0c81a8cde759a533387a.jpg" />

## Description
A desktop graphical user interface application designed to interface directly with a local Microsoft Outlook installation. It synchronizes user-selected mailboxes and facilitates the targeted bulk extraction of email attachments. The interface utilizes the Windows Fluent Design system and dynamically inherits the host operating system's accent color.

---

## System Requirements

| Requirement | Specification |
| :--- | :--- |
| Operating System | Windows (Strictly required for COM/MAPI and Registry access) |
| Local Application | Microsoft Outlook (Installed and authenticated) |
| Runtime Environment | Python 3.8+ |

---

## Dependencies

| Library | Purpose |
| :--- | :--- |
| `PySide6` | Core Qt framework for the main application interface. |
| `qfluentwidgets` | Microsoft Fluent Design UI component implementation. |
| `pywin32` | Windows API integration (`win32com`, `pythoncom`) for Outlook communication. |

**Installation Command:**
```bash
pip install PySide6 qfluentwidgets pywin32
```

---

## Architectural Details

* **Concurrency Management**: Mailbox synchronization executes on a dedicated worker thread (`SyncWorker`). It initializes a separate COM apartment via `pythoncom.CoInitialize()` to safely read Outlook data without freezing the primary GUI event loop. Thread communication is handled via PySide6 `Signal` emissions.
* **Data Retrieval**: The system connects to the Outlook MAPI namespace, maps available accounts and subfolders, sorts items by `ReceivedTime`, and limits the iteration to the 250 most recent items to optimize performance.
* **System Integration**: The script queries the Windows Registry (`HKEY_CURRENT_USER\Software\Microsoft\Windows\DWM`) to extract the system's `AccentColor`, applying a bitwise shift to convert the ABGR integer to an RGB `QColor` for UI consistency.
* **File Sanitization**: Prior to execution of the disk write operation, attachment filenames are processed through a regular expression (`r'[\\/*?:"<>|]'`) to strip invalid characters and prevent OS-level file creation errors.

---

## Execution Configuration

1. Confirm that Microsoft Outlook is configured with at least one active profile on the host machine.
2. Ensure the working directory contains a valid `mail.ico` file to satisfy the system tray and window icon configuration.
3. Launch the application:

```bash
python main.py
```

---

## Operating Procedures

1. Open the application and select the target profile from the **CONTA** dropdown.
2. Select the specific directory to scan from the **PASTA** dropdown.
3. Execute the **Sincronizar** function to begin parsing the mailbox.
4. Select the target attachments via the table checkboxes, or utilize the **Selecionar Todos** function for batch processing.
5. Click **Baixar Selecionados** and specify the absolute path for the extraction output directory.