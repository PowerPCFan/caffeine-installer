import ctypes
import os
# import pythoncom
import requests
import sys
import shutil as sh
import zipfile
from PyQt6.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QCheckBox, QMessageBox, QGroupBox, QProgressBar
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt
from win32com.client import Dispatch
# from win32comext.shell import shell, shellcon
# from win32com.shell import shell, shellcon

# Fallback just in case someone runs the .py file directly or if the .exe fails to elevate itself using uac_admin=True
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    # Relaunch with admin rights
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit()

# THANKS FOR THIS FUNCTION CLAUDE 3.5 SONNET (i couldnt figure out how to get paths to work)
def get_asset(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def new_folder(folder_path):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    return folder_path

def delete_folder(folder_path):
    if os.path.exists(folder_path):
        sh.rmtree(folder_path)

def delete_file(file_path):
    if os.path.exists(file_path):
        os.remove(file_path)
        
class WindowsShortcut:
    def __init__(self, shortcut_path):
        self.shortcut_path = os.path.abspath(shortcut_path)
        self.shell = Dispatch('WScript.Shell')
        self.shortcut = self.shell.CreateShortCut(self.shortcut_path)

    def set_target(self, target_path):
        self.shortcut.Targetpath = os.path.abspath(target_path)

    def set_working_directory(self, working_directory):
        self.shortcut.WorkingDirectory = os.path.abspath(working_directory)

    def set_icon(self, icon_path, icon_index=0):
        self.shortcut.IconLocation = f"{os.path.abspath(icon_path)},{icon_index}"

    def set_arguments(self, arguments):
        self.shortcut.Arguments = arguments

    def set_description(self, description):
        self.shortcut.Description = description

    def save(self):
        self.shortcut.save()

# Vars for app
window_size_x = 480
window_size_y = 350

installer_title = "Caffeine Installer"
installer_description = "This is a simple installer for Caffeine, an app made by Zhorn Software that keeps your PC from falling asleep. This installer allows you to install the 32-bit or 64-bit version of Caffeine, and create shortcuts to it in the Start Menu and on the Desktop."

class InstallerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(installer_title)
        self.setFixedSize(window_size_x, window_size_y)
        self.setWindowIcon(QIcon(get_asset("assets/icon.ico")))

        # Layouts
        main_layout = QVBoxLayout()

        title = QLabel(installer_title)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("font-size: 18pt; font-weight: bold;")
        main_layout.addWidget(title)
        
        description = QLabel(installer_description)
        description.setAlignment(Qt.AlignmentFlag.AlignCenter)
        description.setStyleSheet("font-size: 10pt;")
        description.setWordWrap(True)
        main_layout.addWidget(description)

        # Options group
        options_group = QGroupBox("Options")
        options_layout = QVBoxLayout()
        options_layout.setSpacing(5)

        self.checkbox_32bit = QCheckBox("Install 32-bit version")
        self.checkbox_64bit = QCheckBox("Install 64-bit version")
        self.checkbox_startmenu = QCheckBox("Add Start Menu shortcut")
        self.checkbox_desktop = QCheckBox("Add Desktop shortcut")
                
        # make some options checked by default
        self.checkbox_64bit.setChecked(True)
        self.checkbox_startmenu.setChecked(True)
        self.checkbox_desktop.setChecked(True)

        options_layout.addWidget(self.checkbox_32bit)
        options_layout.addWidget(self.checkbox_64bit)
        options_layout.addWidget(self.checkbox_startmenu)
        options_layout.addWidget(self.checkbox_desktop)

        options_group.setLayout(options_layout)
        main_layout.addWidget(options_group)

        # Install button
        install_btn = QPushButton("Install")
        install_btn.setStyleSheet("font-weight: bold; height: 30px;")
        install_btn.clicked.connect(self.install)
        main_layout.addSpacing(10)
        main_layout.addWidget(install_btn)

        # Progress bar (initially hidden)
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximum(100)
        self.progress_bar.hide()
        main_layout.addWidget(self.progress_bar)

        self.setLayout(main_layout)

    def install(self):
        # Vars
        system_drive = f"{os.environ.get("SYSTEMDRIVE", "C:")}\\"
        temp_dir = os.environ.get("TEMP")
        
        program_files_64 = os.path.join(system_drive, "Program Files")
        program_files_32 = os.path.join(system_drive, "Program Files (x86)")
        
        public_desktop = os.path.join(system_drive, "Users", "Public", "Desktop")
        start_menu = os.path.join(system_drive, 'ProgramData', 'Microsoft', 'Windows', 'Start Menu', 'Programs')
        
        caffeine_url = "https://www.zhornsoftware.co.uk/caffeine/caffeine.zip"

        self.progress_bar.show()
        self.progress_bar.setValue(0)
        
        # make folder to download to
        new_folder(os.path.join(temp_dir, "caffeine-temp"))
        caffeine_zip_save_folder = os.path.join(temp_dir, "caffeine-temp")
        caffeine_zip_save_path = os.path.join(caffeine_zip_save_folder, "caffeine.zip")
        caffeine_binary_folder = os.path.join(caffeine_zip_save_folder, "caffeine")
        # Download file stuff
        try:
            # Download the file
            response = requests.get(caffeine_url)
            self.progress_bar.setValue(5)
            
            if response.status_code == 200:
                # Save the file
                with open(caffeine_zip_save_path, 'wb') as file:
                    file.write(response.content)
                self.progress_bar.setValue(10)
                # Unzip the file
                with zipfile.ZipFile(caffeine_zip_save_path, 'r') as zip_ref:
                    zip_ref.extractall(caffeine_binary_folder)
                # Delete the zip file
                delete_file(caffeine_zip_save_path)
                self.progress_bar.setValue(15)
                # Set the vars for Caffeine's executables
                caffeine64exe = os.path.join(caffeine_binary_folder, "caffeine64.exe")
                caffeine32exe = os.path.join(caffeine_binary_folder, "caffeine32.exe")
                
                downloaded_successfully = True
            else:
                self.progress_bar.setValue(100)
                QMessageBox.critical(self, "Error", f"Error: Failed to download Caffeine. Response code: {response.status_code}")
        except Exception as e:
            self.progress_bar.setValue(100)
            QMessageBox.critical(self, "Error", f"Error: {str(e)}")

        if downloaded_successfully:
            try:
                # 32-bit
                if self.checkbox_32bit.isChecked():
                    caffeine32_install_path = new_folder(os.path.join(program_files_32, "Zhorn Software"))
                    caffeine32_installed_location = os.path.join(caffeine32_install_path, "caffeine32.exe")
                    if os.path.exists(caffeine32exe):
                        sh.copy2(caffeine32exe, caffeine32_installed_location)
                    self.progress_bar.setValue(25)

                # 64-bit
                if self.checkbox_64bit.isChecked():
                    caffeine64_install_path = new_folder(os.path.join(program_files_64, "Zhorn Software"))
                    caffeine64_installed_location = os.path.join(caffeine64_install_path, "caffeine64.exe")
                    if os.path.exists(caffeine64exe):
                        sh.copy2(caffeine64exe, caffeine64_installed_location)
                    self.progress_bar.setValue(50)

                # Create shortcuts
                if self.checkbox_startmenu.isChecked():
                    if self.checkbox_32bit.isChecked():
                        shortcut = WindowsShortcut(f"{start_menu}\\Caffeine (32-bit).lnk")
                        shortcut.set_target(caffeine32_installed_location)
                        shortcut.set_working_directory(caffeine32_install_path)
                        shortcut.set_icon(get_asset("assets/icon.ico"))
                        shortcut.set_description("Caffeine 32-bit")
                        shortcut.save()

                    if self.checkbox_64bit.isChecked():
                        shortcut = WindowsShortcut(f"{start_menu}\\Caffeine.lnk")
                        shortcut.set_target(caffeine64_installed_location)
                        shortcut.set_working_directory(caffeine64_install_path)
                        shortcut.set_icon(get_asset("assets/icon.ico"))
                        shortcut.set_description("Caffeine 64-bit")
                        shortcut.save()
                        
                    self.progress_bar.setValue(75)

                if self.checkbox_desktop.isChecked():
                    if self.checkbox_32bit.isChecked():
                        shortcut = WindowsShortcut(f"{public_desktop}\\Caffeine (32-bit).lnk")
                        shortcut.set_target(caffeine32_installed_location)
                        shortcut.set_working_directory(caffeine32_install_path)
                        shortcut.set_icon(get_asset("assets/icon.ico"))
                        shortcut.set_description("Caffeine 32-bit")
                        shortcut.save()

                    if self.checkbox_64bit.isChecked():
                        shortcut = WindowsShortcut(f"{public_desktop}\\Caffeine.lnk")
                        shortcut.set_target(caffeine64_installed_location)
                        shortcut.set_working_directory(caffeine64_install_path)
                        shortcut.set_icon(get_asset("assets/icon.ico"))
                        shortcut.set_description("Caffeine 64-bit")
                        shortcut.save()
                        
                    self.progress_bar.setValue(90)

                self.progress_bar.setValue(100)
                QMessageBox.information(self, "Success", f"Successfully installed Caffeine!")
            except Exception as e:
                self.progress_bar.setValue(100)
                QMessageBox.critical(self, "Error", f"Error: Failed to install Caffeine: {str(e)}")

        # Clean up
        # Add clean up stuff here
        
# Show the PyQt6 app
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = InstallerApp()
    window.show()
    sys.exit(app.exec())
