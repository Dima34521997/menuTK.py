from PySide6.QtWidgets import (
    QApplication,
    QWidget,
    QPushButton,
    QVBoxLayout,
    QHBoxLayout,
    QFileDialog,
    QLabel,
    QLineEdit,
    QCheckBox,
    QTextBrowser,
)
from PySide6.QtGui import QDesktopServices
from PySide6.QtCore import Qt, QUrl
import sys
import json
import os

# Мои импорты
# перечень элементов, спецификация, ведомость покупная
import PE, S, VP

# Глобальные переменные
files_to_open = []
data = {}

# Путь к Profile.json
PROFILE_FILE = "Profile.json"

# Попытка загрузить Profile.json
try:
    with open(PROFILE_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)
except (FileNotFoundError, json.JSONDecodeError, UnicodeDecodeError):
    data = {
        "Project_Name": "",
        "Templates_Path": "",
        "Razrab": "",
        "Proveril": "",
        "N_control": "",
        "Utverdil": "",
        "Scheme": True,
        "PE": True,
        "DK": True,
        "I1": True,
        "I2": True,
        "SB_count": 1,
        "Precent": 10
    }
    with open(PROFILE_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=4)


class UI(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()
        self.load_data()

    def init_ui(self):
        self.setWindowTitle("Генератор документов")
        self.resize(580, 520)

        # Устанавливаем стили
        self.setStyleSheet("""
            /* Кнопки */
            QPushButton {
                color: white;
                border: none;
                padding: 6px;
                border-radius: 4px;
                font-size: 12px;
                font-weight: 500;
            }
            QPushButton#btn_RemoveFiles {
                background-color: #be6f75;
            }
            QPushButton#btn_RemoveFiles:hover {
                background-color: #a85f65;
            }
            QPushButton#btn_RemoveFiles:pressed {
                background-color: #925056;
            }

            QPushButton#btn_AddFile {
                background-color: #5fa586;
            }
            QPushButton#btn_AddFile:hover {
                background-color: #529073;
            }
            QPushButton#btn_AddFile:pressed {
                background-color: #467a63;
            }

            QPushButton#btn_PE,
            QPushButton#btn_S,
            QPushButton#btn_VP {
                background-color: #548b9c;
            }
            QPushButton#btn_PE:hover,
            QPushButton#btn_S:hover,
            QPushButton#btn_VP:hover {
                background-color: #4a7c8a;
            }
            QPushButton#btn_PE:pressed,
            QPushButton#btn_S:pressed,
            QPushButton#btn_VP:pressed {
                background-color: #3e6a78;
            }

            QPushButton[flat='true'] {
                background-color: #548b9c;
                padding: 4px;
                min-width: 30px;
                font-size: 12px;
            }

            /* Поля ввода */
            QLineEdit {
                padding: 10px 12px;
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                background-color: #ffffff;
                font-size: 14px;
                max-height: 16px;
            }
            QLineEdit:focus {
                border: 1px solid #548b9c;
                background-color: #f8fbfd;
                box-shadow: 0 0 8px rgba(84, 139, 156, 0.2);
            }
            QLineEdit:hover:!focus {
                border: 1px solid #95a5a6;
            }

            /* Метки (labels) */
            QLabel {
                font-size: 15px;
                font-weight: 600;
                color: #2c3e50;
                min-width: 130px;
                text-align: left;
            }

            /* Чекбокс */
            QCheckBox {
                font-size: 14px;
                color: #2c3e50;
                spacing: 8px;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                border: 1px solid #bdc3c7;
                border-radius: 5px;
                background: #ffffff;
            }
            QCheckBox::indicator:checked {
                background: #548b9c;
                border: 1px solid #548b9c;
            }
            QCheckBox::indicator:checked::after {
                content: '✓';
                color: white;
                font-weight: bold;
                text-align: center;
                line-height: 18px;
            }

            /* Текстовое поле (логи) */
            QTextEdit {
                font-family: 'Helvetica', 'Arial', sans-serif;
                font-size: 14px;
                font-weight: 500;
                color: #2c3e50;
                background-color: #f7f1c8;
                border: 1px solid #bdc3c7;
                border-radius: 6px;
                padding: 12px;
            }
        """)

        # Основной layout
        main_layout = QVBoxLayout()

        # --- Кнопки управления файлами ---
        top_layout = QHBoxLayout()
        self.btn_AddFile = QPushButton("Добавить файл/-ы")
        self.btn_RemoveFiles = QPushButton("Очистить")

        self.btn_RemoveFiles.setObjectName("btn_RemoveFiles")
        self.btn_AddFile.setObjectName("btn_AddFile")

        top_layout.addWidget(self.btn_AddFile)
        top_layout.addWidget(self.btn_RemoveFiles)
        main_layout.addLayout(top_layout)

        # --- Список файлов (логи) ---
        self.txt_FilesListView = QTextBrowser()
        self.txt_FilesListView.setOpenExternalLinks(False)  # Отключаем автоматическое открытие
        self.txt_FilesListView.setMaximumHeight(180)
        main_layout.addWidget(self.txt_FilesListView)

        # Подключаем обработчик ссылок
        self.txt_FilesListView.anchorClicked.connect(self.on_anchor_clicked)

        # --- Кнопки генерации ---
        button_layout = QHBoxLayout()
        self.btn_PE = QPushButton("Перечень")
        self.btn_S = QPushButton("Спецификация")
        self.btn_VP = QPushButton("ВП")

        self.btn_PE.setObjectName("btn_PE")
        self.btn_S.setObjectName("btn_S")
        self.btn_VP.setObjectName("btn_VP")

        button_layout.addWidget(self.btn_PE)
        button_layout.addWidget(self.btn_S)
        button_layout.addWidget(self.btn_VP)
        main_layout.addLayout(button_layout)

        # --- Форма с полями ---
        form_layout = QVBoxLayout()

        # Вспомогательная функция для создания строки "label + input + optional button"
        def add_field(layout, label_text, line_edit, button=None):
            h_layout = QHBoxLayout()
            label = QLabel(label_text)
            label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            h_layout.addWidget(label)
            h_layout.addWidget(line_edit)
            if button:
                h_layout.addWidget(button)
            layout.addLayout(h_layout)

        # --- Поля ---
        self.ent_Templates = QLineEdit()
        self.btn_SelectTemplate = QPushButton("Выбрать...")
        self.btn_SelectTemplate.setFixedSize(66, 36)
        self.btn_SelectTemplate.setProperty("flat", "true")
        self.btn_SelectTemplate.clicked.connect(self.on_select_template_folder)
        add_field(form_layout, "Путь к шаблонам", self.ent_Templates, self.btn_SelectTemplate)

        self.ent_Project = QLineEdit()
        add_field(form_layout, "Название проекта", self.ent_Project)

        self.ent_Razrab = QLineEdit()
        add_field(form_layout, "Разработчик", self.ent_Razrab)

        self.ent_Checked = QLineEdit()
        add_field(form_layout, "Проверил", self.ent_Checked)

        self.ent_Control = QLineEdit()
        add_field(form_layout, "Н. контроль", self.ent_Control)

        self.ent_Utverdil = QLineEdit()
        add_field(form_layout, "Утвердил", self.ent_Utverdil)

        # --- Чекбокс ---
        self.c1 = QCheckBox("Гражданский")
        form_layout.addWidget(self.c1, alignment=Qt.AlignLeft)

        main_layout.addLayout(form_layout)
        self.setLayout(main_layout)

        # --- Подключение сигналов ---
        self.btn_AddFile.clicked.connect(self.onOpen)
        self.btn_RemoveFiles.clicked.connect(self.clearFiles)
        self.btn_PE.clicked.connect(self.create_pe)
        self.btn_S.clicked.connect(self.create_s)
        self.btn_VP.clicked.connect(self.create_vp)

        self.ent_Templates.textChanged.connect(self.on_modified)
        self.ent_Project.textChanged.connect(self.on_modified)
        self.ent_Razrab.textChanged.connect(self.on_modified)
        self.ent_Checked.textChanged.connect(self.on_modified)
        self.ent_Control.textChanged.connect(self.on_modified)
        self.ent_Utverdil.textChanged.connect(self.on_modified)

    def on_anchor_clicked(self, url):
        """Обработка клика по ссылке в логе"""
        print(f"[LOG] Переход по ссылке: {url.toString()}")
        QDesktopServices.openUrl(url)
        self.txt_FilesListView.ensureCursorVisible()  # Убедимся, что видно последнюю строку

    def load_data(self):
        self.ent_Templates.setText(data.get("Templates_Path", ""))
        self.ent_Project.setText(data.get("Project_Name", ""))
        self.ent_Razrab.setText(data.get("Razrab", ""))
        self.ent_Checked.setText(data.get("Proveril", ""))
        self.ent_Control.setText(data.get("N_control", ""))
        self.ent_Utverdil.setText(data.get("Utverdil", ""))

    def on_select_template_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выберите папку с шаблонами", self.ent_Templates.text())
        if folder:
            folder = os.path.normpath(folder)
            self.ent_Templates.setText(folder)
            data['Templates_Path'] = folder
            self.save_profile()

    def on_modified(self):
        data['Templates_Path'] = self.ent_Templates.text()
        data['Project_Name'] = self.ent_Project.text()
        data['Razrab'] = self.ent_Razrab.text()
        data['Proveril'] = self.ent_Checked.text()
        data['N_control'] = self.ent_Control.text()
        data['Utverdil'] = self.ent_Utverdil.text()
        self.save_profile()

    def save_profile(self):
        try:
            with open(PROFILE_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            self.log_error(f"Не удалось сохранить Profile.json: {e}")

    def onOpen(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Выберите файлы", "", "Excel files (*.xlsx *.xls)"
        )
        if files:
            files_to_open.clear()
            self.txt_FilesListView.clear()
            for file in files:
                files_to_open.append(file)
                self.txt_FilesListView.insertHtml(f"{os.path.basename(file)}<br>")
            self.txt_FilesListView.ensureCursorVisible()

    def clearFiles(self):
        files_to_open.clear()
        self.txt_FilesListView.clear()

    def log_error(self, message):
        html = f'<span style="color: red; font-size: 11pt;">{message}</span><br>'
        self.txt_FilesListView.insertHtml(html)
        self.txt_FilesListView.ensureCursorVisible()

    def log_success(self, message):
        html = f'<span style="color: green; font-size: 11pt;">{message}</span><br>'
        self.txt_FilesListView.insertHtml(html)
        self.txt_FilesListView.ensureCursorVisible()

    def create_pe(self):
        """Создать перечень элементов"""
        if not files_to_open:
            self.log_error("Вы не выбрали ни одного файла!")
            return

        try:
            created_files = PE.execute(files_to_open, self.c1.isChecked())
            self.log_success("ПЭ успешно создан(-ы)!")

            for file_path in created_files:
                file_name = os.path.basename(file_path)
                folder_path = os.path.dirname(file_path)
                folder_url = QUrl.fromLocalFile(folder_path).toString()

                link = (
                    f'<a href="{folder_url}" '
                    f'style="color: #0066cc; text-decoration: underline; font-size: 14px;">'
                    f'📁 Открыть папку с {file_name}</a><br>'
                )
                self.txt_FilesListView.insertHtml(link)
            self.txt_FilesListView.ensureCursorVisible()

        except IOError:
            self.log_error("Ошибка при создании ПЭ! Закройте открытые экземпляры перечня!")
        except Exception as e:
            self.log_error(f"Неизвестная ошибка: {e}")

    def create_s(self):
        """Создать спецификацию"""
        if not files_to_open:
            self.log_error("Не выбрано ни одного файла!")
            return
        try:
            S.execute(files_to_open, self.c1.isChecked())
            self.log_success("СП создана(-ы)!")
        except Exception as e:
            self.log_error(f"Ошибка при создании СП: {e}")

    def create_vp(self):
        """Создать ведомость покупную"""
        if not files_to_open:
            self.log_error("Не выбрано ни одного файла!")
            return
        try:
            VP.execute(files_to_open)
            self.log_success("ВП создана!")
        except Exception as e:
            self.log_error(f"Ошибка при создании ВП: {e}")


def main():
    app = QApplication(sys.argv)
    ui = UI()
    ui.show()

    def save_on_exit():
        data['Templates_Path'] = ui.ent_Templates.text()
        data['Project_Name'] = ui.ent_Project.text()
        data['Razrab'] = ui.ent_Razrab.text()
        data['Proveril'] = ui.ent_Checked.text()
        data['N_control'] = ui.ent_Control.text()
        data['Utverdil'] = ui.ent_Utverdil.text()
        try:
            with open(PROFILE_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
        except Exception as e:
            print(f"Ошибка сохранения при выходе: {e}")

    app.aboutToQuit.connect(save_on_exit)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()