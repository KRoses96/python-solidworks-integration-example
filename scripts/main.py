"""
Author: Manuel Rosa
Description: Main GUI to run all individual executables and change settings

"""

import sys
import subprocess
from PyQt5.QtWidgets import (
    QApplication,
    QMainWindow,
    QLabel,
    QPushButton,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QFrame,
    QMessageBox,
    QLineEdit,
    QCheckBox,
    QDialog,
    QFileDialog,
    QSlider,
)
from PyQt5.QtGui import QPixmap, QColor, QCursor, QGuiApplication
from PyQt5.QtCore import Qt, QTimer
import os
import shutil

base_width = 1920
script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
parent_dir = os.path.dirname(script_dir)
options = os.path.join(parent_dir, "extras", "op.txt")

with open(options, "r") as file:
    lines = file.readlines()

aspect = lines[5].split("=")[1].strip() if len(lines) >= 6 and "=" in lines[5] else None

with open(options, "r") as file:
    lines = file.readlines()

if len(lines) >= 3:
    lines[2] = f"Path = {script_dir}\n"

with open(options, "w") as file:
    file.writelines(lines)


class SettingsDialog(QDialog):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Definições")
        screen = QGuiApplication.primaryScreen()
        screen_size = screen.availableSize()
        aspect_fix = (screen_size.width() / base_width) * float(aspect)
        Geo1 = int(70 * aspect_fix)
        Geo2 = int(70 * aspect_fix)
        Geo3 = int(170 * aspect_fix)
        Geo4 = int(70 * aspect_fix)
        self.setGeometry(Geo1, Geo2, Geo3, Geo4)
        self.setWindowFlag(Qt.WindowContextHelpButtonHint, False)
        self.setStyleSheet("background-color: #34495E;color: #fff")
        with open(options, "r") as file:
            lines = file.readlines()
        if lines:
            first_line = lines[
                3
            ].strip()  
            if first_line:
                optimi = first_line[-1]  
        self.optimi = optimi 

        layout = QVBoxLayout()

        set_solidworks_path_button = QPushButton("Definir Caminho Solidworks")
        set_solidworks_path_button.clicked.connect(self.set_solidworks_path)
        set_solidworks_path_button.setStyleSheet("background-color: #A00000")
        layout.addWidget(set_solidworks_path_button)

        set_phc_path_button = QPushButton("Definir Caminho Base Dados")
        set_phc_path_button.clicked.connect(self.set_phc_path)
        set_phc_path_button.setStyleSheet("background-color: #A00000")
        layout.addWidget(set_phc_path_button)

        perfil_button = QPushButton("Regras de Aproveitamento para Perfis")
        perfil_button.clicked.connect(self.run_perfil)
        perfil_button.setStyleSheet("background-color: #A00000")

        self.toggle_checkbox = QCheckBox("Aproveitamento de Perfis")
        self.toggle_checkbox.setChecked(
            self.optimi == "1"
        )  
        self.toggle_checkbox.stateChanged.connect(self.toggle_optimi)

        sheet_button_1 = QPushButton("Chapas / Nesting")
        sheet_button_1.clicked.connect(self.run_chapa_nest)
        sheet_button_1.setStyleSheet("background-color: #A00000")

        sheet_button_2 = QPushButton("Lista Chapas Manual")
        sheet_button_2.clicked.connect(self.run_m_sheet_editor)
        sheet_button_2.setStyleSheet("background-color: #A00000")
        layout.addWidget(sheet_button_2)

        self.aspect_slider = QSlider(Qt.Horizontal)
        self.aspect_slider.setMinimum(50)  
        self.aspect_slider.setMaximum(200)  
        self.aspect_slider.setSliderPosition(100)  
        self.aspect_slider.valueChanged.connect(self.update_aspect_label)

        self.aspect_label = QLabel(f"Tamanho Janela: {float(aspect)*100} %")
        apply_button = QPushButton("Guardar Tamanho Janela", self)
        apply_button.clicked.connect(self.accept)
        apply_button.setStyleSheet("background-color: #A00000")

        separator_row = QHBoxLayout()
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Raised)
        separator.setStyleSheet("background-color: #FFFFFF")  
        separator.setFixedHeight(2)  
        separator_row.addWidget(separator)
        manual_button = QPushButton("Manual de Utilização", self)
        manual_button.clicked.connect(self.manual_open)
        font = manual_button.font()
        font.setPointSize(12)  
        manual_button.setFont(font)
        manual_button.setStyleSheet(
            "QPushButton {"
            "   background-color: grey;"
            "   border-radius: 10px;"  
            "   padding: 10px;"  
            "   color: white;"  
            "}"
            "QPushButton:hover {"
            "   background-color: darkgrey;"  
            "}"
        )

        layout.addWidget(self.aspect_label)
        layout.addWidget(self.aspect_slider)
        layout.addWidget(apply_button)
        layout.addLayout(separator_row)
        layout.addWidget(manual_button)
        self.setLayout(layout)

    def update_aspect_label(self):
        aspect_value = self.aspect_slider.value()
        self.aspect_label.setText(f"Tamanho Janela: {aspect_value:.2f} %")

    def get_aspect_value(self):
        return aspect / 100.0

    def manual_open(self):
        exe = "Manual.pdf"
        exe_to_run = os.path.join(parent_dir, "extras", exe)
        file_path = exe_to_run  
        subprocess.Popen(file_path, shell=True)

    def accept(self):
        aspect_value = self.aspect_slider.value() / 100.0
        if aspect_value == 1.0:
            sys.exit(0)
        else:

            with open(options, "r") as file:
                lines = file.readlines()

            lines[5] = f"aspect = {aspect_value}\n"

            with open(options, "w") as file:
                file.writelines(lines)
            self.restart()
            sys.exit(0)

    def restart(self):
        exe = "main.exe"
        exe_to_run = os.path.join(parent_dir, "exe", exe)
        file_path = exe_to_run 
        subprocess.Popen(file_path, shell=True)

    def set_solidworks_path(self):
        options = QFileDialog.Options()
        options |= QFileDialog.ReadOnly
        file_filter = "Executable Files (*.exe)"
        path_choosen_by_user, _ = QFileDialog.getOpenFileName(
            self, "Selecione SLDWORKS.exe", "", file_filter, options=options
        )
        if path_choosen_by_user:
            self.update_solidworks_path_in_file(path_choosen_by_user)

    def set_phc_path(self):
        path_choosen_by_user = QFileDialog.getExistingDirectory(
            self, "Selecione a pasta IMP na pasta Produção"
        )
        if path_choosen_by_user:
            self.update_phc_path_in_file(path_choosen_by_user)

    def update_solidworks_path_in_file(self, path_choosen_by_user):
        with open(options, "r") as file:
            lines = file.readlines()
        if lines:
            lines[1] = (
                f"Path = {path_choosen_by_user}\n"  
            )
        with open(options, "w") as file:
            file.writelines(lines)

    def update_phc_path_in_file(self, path_choosen_by_user):
        with open(options, "r") as file:
            lines = file.readlines()
        if lines:
            lines[9] = f"phc = {path_choosen_by_user}\n"
        with open(options, "w") as file:
            file.writelines(lines)

    def run_perfil(self):
        exe = "csv_editor.exe"
        exe_to_run = os.path.join(parent_dir, "exe", exe)
        file_path = exe_to_run  
        subprocess.Popen(file_path, shell=True)

    def run_chapa_nest(self):
        exe = "csv_corte_editor.exe"
        exe_to_run = os.path.join(parent_dir, "exe", exe)
        file_path = exe_to_run  
        subprocess.Popen(file_path, shell=True)

    def run_m_sheet_editor(self):
        exe = "m_sheet_editor.exe"
        exe_to_run = os.path.join(parent_dir, "exe", exe)
        file_path = exe_to_run  
        subprocess.Popen(file_path, shell=True)

    def toggle_optimi(self, state):
        if state == Qt.Checked:
            self.optimi = "1"
        else:
            self.optimi = "0"
        self.update_optimi_in_file()

    def update_optimi_in_file(self):
        with open(options, "r") as file:
            lines = file.readlines()
        if lines:
            lines[3] = f"optimi = {self.optimi}\n"  
        with open(options, "w") as file:
            file.writelines(lines)

    def read_optimi_value(self):
        with open(options, "r") as file:
            lines = file.readlines()
            if lines:
                for line in lines:
                    if line.startswith("optimi"):
                        return line.strip().split("=")[-1].strip()
        return "0"


class CustomTitleBar(QWidget):
    def __init__(self, settings_dialog):
        super().__init__()
        self.init_ui()
        self.settings_dialog = settings_dialog
        self.dragging = False
        self.offset = None

    def init_ui(self):
        screen = QGuiApplication.primaryScreen()
        screen_size = screen.availableSize()
        aspect_fix = (screen_size.width() / base_width) * float(aspect)
        self.setStyleSheet("background-color: #2E3D4E; color: #fff; padding: 5px;")
        self.setStyleSheet(
            """
            QWidget {
                background-color: #2E3D4E;
                border-radius: 12px;
            }
        """
        )
        self.setFixedHeight(int(70 * aspect_fix))
        layout = QVBoxLayout()
        title_bar_row = QHBoxLayout()
        setting_button = QPushButton("⚙")
        setting_button.clicked.connect(self.setting_window)
        setting_button.setStyleSheet(
            """
            background-color: #2E3D4E;
            color: #fff;
            border: none;
            padding: 4px;
            font-size: {}pt;
        """.format(
                10 * aspect_fix
            )
        )
        title_bar_row.addWidget(setting_button)
        title_label = QLabel("Projecto")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setStyleSheet(
            "font-size: {}pt;font-weight: bold; color: #fff;".format(10 * aspect_fix)
        )
        title_bar_row.addWidget(title_label, 1)
        title_bar_row.addStretch(0)
        minimize_button = QPushButton("—")
        minimize_button.clicked.connect(self.minimize_window)
        minimize_button.setStyleSheet(
            """
            background-color: #2E3D4E;
            color: #fff;
            border: none;
            padding: 5px;
            font-size: {}pt;
        """.format(
                10 * aspect_fix
            )
        )
        title_bar_row.addWidget(minimize_button)
        close_button = QPushButton("✕")
        close_button.clicked.connect(self.close_window)
        close_button.setStyleSheet(
            """
            background-color: #E74C3C;
            color: #fff;
            border: none;
            padding: 5px;
            font-size: {}pt;
        """.format(
                10 * aspect_fix
            )
        )
        title_bar_row.addWidget(close_button)
        layout.addLayout(title_bar_row)
        separator_row = QHBoxLayout()
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Raised)
        separator.setStyleSheet("background-color: #FFFFFF")  
        separator.setFixedHeight(1)  
        separator_row.addWidget(separator)
        layout.addLayout(separator_row)
        self.setLayout(layout)

    def setting_window(self):
        main_win_geometry = self.parent().geometry()
        settings_x = main_win_geometry.x() + main_win_geometry.width()
        settings_y = main_win_geometry.y()
        self.settings_dialog.move(settings_x, settings_y)
        self.settings_dialog.show()

    def minimize_window(self):
        self.parent().showMinimized()

    def close_window(self):
        sys.exit(0)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.dragging = True
            self.offset = event.globalPos() - self.parent().pos()

    def mouseMoveEvent(self, event):
        if self.dragging:
            self.parent().move(event.globalPos() - self.offset)

    def mouseReleaseEvent(self, event):
        self.dragging = False


class CompanyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setGeometry(50, 50, 0, 0)  
        self.setWindowTitle("Projeto")
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setMenuWidget(CustomTitleBar(settings_dialog))
        self.setStyleSheet(
            """
            QMainWindow {
                background-color: #34495E;
                border: 1px solid #2C3E50;
                border-radius: 12px;
            }
        """
        )
        self.init_ui()

    def sleep5sec(self):
        for button in self.findChildren(QPushButton):
            button.setEnabled(False)
        QTimer.singleShot(1000, self.enable_all_buttons)

    def enable_all_buttons(self):
        for button in self.findChildren(QPushButton):
            button.setEnabled(True)

    def init_ui(self):
        screen = QGuiApplication.primaryScreen()
        screen_size = screen.availableSize()
        logo_path = os.path.join(parent_dir, "extras", "Logo_nb.png")
        aspect_fix = (screen_size.width() / base_width) * float(aspect)
        pixmap = QPixmap(logo_path)
        pixmap = pixmap.scaledToHeight(
            int(80 * (aspect_fix * 1.7)), Qt.SmoothTransformation
        ) 
        logo_label = QLabel(self)
        logo_label.setPixmap(pixmap)
        logo_label.setAlignment(Qt.AlignCenter)

        self.central_widget = QWidget()
        self.setCentralWidget(self.central_widget)
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        separator.setFixedHeight(1)
        separator.setStyleSheet("background-color: #FFFFFF;border: none")
        separator_list = QFrame()
        separator_list.setFrameShape(QFrame.HLine)
        separator_list.setFrameShadow(QFrame.Sunken)
        separator_list.setFixedHeight(1)
        separator_list.setStyleSheet(" background-color: #FFFFFF;border: none")

        separator_bg = QFrame()
        separator_bg.setFrameShape(QFrame.HLine)
        separator_bg.setFrameShadow(QFrame.Sunken)
        separator_bg.setFixedHeight(1)
        separator_bg.setStyleSheet("background-color: #FFFFFF;border: none")

        separator_prod = QFrame()
        separator_prod.setFrameShape(QFrame.HLine)
        separator_prod.setFrameShadow(QFrame.Sunken)
        separator_prod.setFixedHeight(1)
        separator_prod.setStyleSheet("background-color: #FFFFFF;border: none")

        layout = QVBoxLayout()
        layout.addWidget(logo_label)
        layout.addWidget(separator)  
        macrorun_path = os.path.join(parent_dir, "exe", "macrorun.exe")
        imp_2_path = os.path.join(parent_dir, "exe", "Imp_pdf.exe")
        dxfcord_path = os.path.join(parent_dir, "exe", "dxfcord.exe")
        data_pass_path = os.path.join(parent_dir, "exe", "data_pass.exe")
        sections = [
            (
                "Listas",
                [
                    ("Todas as Listas", macrorun_path),
                    ("Perfis", macrorun_path),
                    ("Chapas", macrorun_path),
                    ("Material", macrorun_path),
                ],
            ),
            (
                "Desenhos",
                [("DWG -> PDF -> Lista", macrorun_path), ("PDF -> Lista", imp_2_path)],
            ),
            ("Maquinas", [("Furação", dxfcord_path)]),
            (
                "Produção",
                [  
                    ("Enviar Projeto", data_pass_path),
                ],
            ),
        ]
        print(aspect_fix)
        less_vibrant_red = QColor(160, 0, 0)
        font_size = 10 * aspect_fix

        def set_button_hover_style(button):
            normal = """
                background-color: {};
                color: white;
                border: 0.5px solid white;
                padding: 7px;
                font-size: {}pt;
                border-radius: 5px;
            """.format(
                less_vibrant_red.name(), font_size
            )
            pressed = """
                background-color: {};
                color: white;
                border: 0.5px solid white;
                padding: 7px;
                font-size: {}pt;
                border-radius: 5px;
            """.format(
                less_vibrant_red.darker(140).name(), font_size
            )
            button.setStyleSheet(normal)
            button.setCursor(QCursor(Qt.PointingHandCursor))

            def set_button_state(state):
                if state == Qt.Unchecked:
                    button.setStyleSheet(normal)
                elif state == Qt.PartiallyChecked:
                    button.setStyleSheet(pressed)
                    timer = QTimer(self)
                    timer.timeout.connect(lambda: button.setStyleSheet(normal))
                    timer.start(250)

            button.setCheckable(True)
            button.setAutoExclusive(True)
            button.toggled.connect(set_button_state)

        for section_name, buttons in sections:
            if section_name:  
                section_label = QLabel(section_name)
                section_label.setAlignment(Qt.AlignCenter)
                section_label.setStyleSheet(
                    "font-size: {}pt; font-weight: bold; color: #fff;".format(
                        13 * aspect_fix
                    )
                )
                layout.addWidget(section_label)
            if section_name == "Listas":
                v_layout = QVBoxLayout()
                for text, path in buttons:
                    button = QPushButton(text)
                    button.clicked.connect(
                        lambda checked, p=path: self.run_executable(p)
                    )
                    set_button_hover_style(button)
                    v_layout.addWidget(button)
                layout.addLayout(v_layout)
                layout.addWidget(separator_list)
            elif section_name == "Desenhos":
                v_layout = QVBoxLayout()
                for text, path in buttons:
                    button = QPushButton(text)
                    button.clicked.connect(
                        lambda checked, p=path: self.run_executable(p)
                    )
                    set_button_hover_style(button)
                    v_layout.addWidget(button)
                layout.addLayout(v_layout)
                layout.addWidget(separator_bg)
            elif section_name == "Maquinas":
                v_layout = QVBoxLayout()
                for text, path in buttons:
                    button = QPushButton(text)
                    button.clicked.connect(
                        lambda checked, p=path: self.run_executable(p)
                    )
                    set_button_hover_style(button)
                    v_layout.addWidget(button)
                layout.addLayout(v_layout)
                layout.addWidget(separator_prod)
            elif section_name == "Produção":
                v_layout = QVBoxLayout()
                for text, path in buttons:
                    button = QPushButton(text)
                    button.clicked.connect(
                        lambda checked, p=path: self.run_executable(p)
                    )
                    set_button_hover_style(button)
                    v_layout.addWidget(button)
                layout.addLayout(v_layout)
        self.central_widget.setLayout(layout)

    def run_executable(self, file_path):
        try:
            self.sleep5sec()
            if self.sender().text() == "Todas as Listas":
                print("Changing op.txt...")
                with open(options, "r+") as file:
                    try:
                        folder_path = r"C:\ProgramData\DXFs_Macro"
                        shutil.rmtree(folder_path)
                    except:
                        pass
                    lines = file.readlines()
                    if lines:
                        lines[0] = lines[0][:-2] + "0\n"
                        lines[7] = lines[7][:-2] + "0\n"
                        file.seek(0)
                        file.writelines(lines)
                        file.truncate()
                        print("op.txt changed to:", lines[0])
            if self.sender().text() == "Perfis":
                print("Changing op.txt...")
                with open(options, "r+") as file:
                    lines = file.readlines()
                    if lines:
                        lines[0] = lines[0][:-2] + "1\n"
                        file.seek(0)
                        file.writelines(lines)
                        file.truncate()
                        print("op.txt changed to:", lines[0])
            if self.sender().text() == "Chapas":
                print("Changing op.txt...")
                try:
                    folder_path = r"C:\ProgramData\DXFs_Macro"
                    shutil.rmtree(folder_path)
                except:
                    pass
                with open(options, "r+") as file:
                    lines = file.readlines()
                    if lines:
                        lines[0] = lines[0][:-2] + "2\n"
                        lines[7] = lines[7][:-2] + "0\n"
                        file.seek(0)
                        file.writelines(lines)
                        file.truncate()
                        print("op.txt changed to:", lines[0])
            if self.sender().text() == "Material":
                print("Changing op.txt...")
                with open(options, "r+") as file:
                    lines = file.readlines()
                    if lines:
                        lines[0] = lines[0][:-2] + "3\n"
                        file.seek(0)
                        file.writelines(lines)
                        file.truncate()
                        print("op.txt changed to:", lines[0])
            if self.sender().text() == "DWG -> PDF -> Lista":
                print("changing op.txt...")
                with open(options, "r+") as file:
                    lines = file.readlines()
                    if lines:
                        lines[0] = lines[0][:-2] + "4\n"
                        file.seek(0)
                        file.writelines(lines)
                        file.truncate()
                        print("op.txt changed to:", lines[0])
                        QMessageBox.about(
                            self,
                            "Alerta!",
                            "Escolha o Assembly correspondente aos desenhos!",
                        )
            print("Running executable:", file_path)
            subprocess.Popen(file_path, shell=True)

        except Exception as e:
            print(f"Error: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    settings_dialog = SettingsDialog()
    mainWin = CompanyApp()
    mainWin.show()
    sys.exit(app.exec_())
