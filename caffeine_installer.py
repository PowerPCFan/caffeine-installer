import ctypes
import os
import psutil
import requests
import sys
import shutil as sh
import time
import threading
import zipfile
from PyQt6.QtWidgets import QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QCheckBox, QMessageBox, QGroupBox, QProgressBar
from PyQt6.QtGui import QIcon
from PyQt6.QtCore import Qt, QTimer
from win32com.client import Dispatch

# Fallback just in case someone runs the .py file directly or if the .exe fails to elevate itself using uac_admin=True
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    # Relaunch with admin rights
    ctypes.windll.shell32.ShellExecuteW(
        None, "runas", sys.executable, " ".join(sys.argv), None, 1
    )
    sys.exit()

# THANKS FOR THIS FUNCTION CLAUDE 3.5 SONNET (i couldnt figure out how to get paths to work)
def get_asset(relative_path):
    try:
        base_path = sys._MEIPASS # type: ignore
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def new_folder(folder_path: str):
    if not os.path.exists(folder_path):
        os.makedirs(folder_path)
    return str(folder_path)

def remove_folder(folder_path: str):
    if os.path.exists(folder_path):
        sh.rmtree(folder_path)

def remove_file(file_path: str):
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

# Claude 3.7 Sonnet gave me the idea: https://claude.ai/chat/2f1c2c3e-9ea1-437a-985d-fe193fdfe177
# but this has been heavily modified
class ProcessManager:
    def __init__(self, process_path):
        self.process_path = process_path
        
    def GetPid(self):
        pid_list = []
        for proc in psutil.process_iter(['name']):
            try:
                if proc.info['name'].lower() == self.process_path.lower():
                    pid_list.append(proc.pid)
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                pass
        return pid_list
    
    @staticmethod
    def Running(pid):
        # Validate PID
        if not isinstance(pid, int) or pid <= 0:
            # Invalid PID, stop function
            return
        
        try:            
            if not psutil.pid_exists(pid):
                # Returns False if the PID doesn't exist
                return False
                
            process = psutil.Process(pid)
            # Returns True if the process is running and not a "zombie"
            # Otherwise, returns False
            return process.is_running() and process.status() != psutil.STATUS_ZOMBIE
            
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # NoSuchProcess: Process doesn't exist
            # AccessDenied: Process exists but no permission to access it
            # ZombieProcess: Process is a zombie (terminated but not reaped)
            return False
    
    @staticmethod
    def Kill(pid, timeout=5):
        """
        Kills a process.
        """
            
        try:
            process = psutil.Process(pid)
            
            # Try graceful termination first
            process.terminate()
            
            # Wait for process to term
            try:
                process.wait(timeout=timeout)
                return True
            except psutil.TimeoutExpired:
                # Force kill if graceful termination times out
                process.kill()
                process.wait(timeout=2)
                return True
                
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            # Process already terminated or no permission to terminate
            pass

# CONSTANTS
system_drive = f"{os.environ.get("SYSTEMDRIVE", "C:")}\\"
user_profile = os.environ.get("USERPROFILE", os.path.expanduser("~"))
temp_dir = os.environ.get("TEMP", os.path.join(user_profile, "AppData", "Local", "Temp"))
program_files_64 = os.path.join(system_drive, "Program Files")
program_files_32 = os.path.join(system_drive, "Program Files (x86)")
public_desktop = os.path.join(system_drive, "Users", "Public", "Desktop")
start_menu = os.path.join(system_drive, 'ProgramData', 'Microsoft', 'Windows', 'Start Menu', 'Programs')
caffeine_url = "https://www.zhornsoftware.co.uk/caffeine/caffeine.zip"
caffeine32_install_path = new_folder(os.path.join(program_files_32, "Zhorn Software"))
caffeine64_install_path = new_folder(os.path.join(program_files_64, "Zhorn Software"))

window_size_x = 480
installer_title = "Caffeine Installer"
installer_description = "This is a simple installer for Caffeine, an app made by Zhorn Software that keeps your PC from falling asleep. This installer allows you to install the 32-bit or 64-bit version of Caffeine, and create shortcuts to it in the Start Menu and on the Desktop."

class InstallerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(installer_title)
        self.setFixedWidth(window_size_x)
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
        options_group = QGroupBox("Installation Options")
        options_layout = QVBoxLayout()
        options_layout.setSpacing(4)

        self.checkbox_32bit = QCheckBox("Install 32-bit version")
        self.checkbox_64bit = QCheckBox("Install 64-bit version")
        self.checkbox_startmenu = QCheckBox("Add Start Menu shortcut")
        self.checkbox_desktop = QCheckBox("Add Desktop shortcut")
        
        options_layout.addWidget(self.checkbox_32bit)
        options_layout.addWidget(self.checkbox_64bit)
        options_layout.addWidget(self.checkbox_startmenu)
        options_layout.addWidget(self.checkbox_desktop)
        
        # make some options checked by default
        self.checkbox_64bit.setChecked(True)
        self.checkbox_startmenu.setChecked(True)
        self.checkbox_desktop.setChecked(True)

        options_group.setLayout(options_layout)
        main_layout.addWidget(options_group)
        
        buttons_group = QWidget()
        buttons_layout = QVBoxLayout()
        buttons_layout.setSpacing(4)

        # Install button
        install_btn = QPushButton("Install")
        install_btn.setStyleSheet("height: 25px;")
        install_btn.clicked.connect(self.install)
        main_layout.addWidget(install_btn)
        
        # Uninstall button
        uninstall_btn = QPushButton("Uninstall")
        uninstall_btn.setStyleSheet("height: 25px;")
        uninstall_btn.clicked.connect(self.start_uninstall)
        main_layout.addWidget(uninstall_btn)

        buttons_group.setLayout(buttons_layout)
        main_layout.addWidget(buttons_group)

        # Progress bar (initially hidden)
        self.progress_bar = QProgressBar()
        self.progress_bar.setMaximum(100)
        self.progress_bar.hide()
        main_layout.addWidget(self.progress_bar)

        main_layout.addSpacing(10)
        self.setLayout(main_layout)
        
    def disable_buttons(self):
        for button in self.findChildren(QPushButton):
            button.setEnabled(False)
        for checkbox in self.findChildren(QCheckBox):
            checkbox.setEnabled(False)

    def install(self):
        self.disable_buttons()
        
        # Initialize some variables
        caffeine32_installed_location: str = ""
        caffeine64_installed_location: str = ""
        caffeine32_copied_icon: str = ""
        caffeine64_copied_icon: str = ""
        caffeine32exe: str = ""
        caffeine64exe: str = ""
        
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
                remove_file(caffeine_zip_save_path)
                self.progress_bar.setValue(15)
                # Set the vars for Caffeine's executables
                caffeine64exe = os.path.join(caffeine_binary_folder, "caffeine64.exe")
                caffeine32exe = os.path.join(caffeine_binary_folder, "caffeine32.exe")
                
                downloaded_successfully = True
            else:
                downloaded_successfully = False
                self.progress_bar.hide()
                QMessageBox.critical(self, "Error", f"Error: Failed to download Caffeine. Response code: {response.status_code}")
        except Exception as e:
            downloaded_successfully = False
            self.progress_bar.hide()
            QMessageBox.critical(self, "Error", f"Error: Failed to download Caffeine: {str(e)}")

        if downloaded_successfully:
            try:
                # 32-bit
                if self.checkbox_32bit.isChecked():
                    caffeine32_installed_location = os.path.join(caffeine32_install_path, "caffeine32.exe")
                    if os.path.exists(caffeine32exe):
                        sh.copy2(caffeine32exe, caffeine32_installed_location)
                    caffeine32_copied_icon = os.path.join(caffeine32_install_path, "icon.ico")
                    if os.path.exists(get_asset("assets/icon.ico")):
                        sh.copy2(get_asset("assets/icon.ico"), caffeine32_copied_icon)
                    self.progress_bar.setValue(25)

                # 64-bit
                if self.checkbox_64bit.isChecked():
                    caffeine64_installed_location = os.path.join(caffeine64_install_path, "caffeine64.exe")
                    if os.path.exists(caffeine64exe):
                        sh.copy2(caffeine64exe, caffeine64_installed_location)
                    caffeine64_copied_icon = os.path.join(caffeine64_install_path, "icon.ico")
                    if os.path.exists(get_asset("assets/icon.ico")):
                        sh.copy2(get_asset("assets/icon.ico"), caffeine64_copied_icon)
                    self.progress_bar.setValue(50)

                # Create shortcuts
                if self.checkbox_startmenu.isChecked():
                    if self.checkbox_32bit.isChecked():
                        start_menu_shortcut_32bit = WindowsShortcut(f"{start_menu}\\Caffeine (32-bit).lnk")
                        start_menu_shortcut_32bit.set_target(caffeine32_installed_location)
                        start_menu_shortcut_32bit.set_working_directory(caffeine32_install_path)
                        start_menu_shortcut_32bit.set_icon(caffeine32_copied_icon)
                        start_menu_shortcut_32bit.set_description("Caffeine 32-bit")
                        start_menu_shortcut_32bit.save()

                    if self.checkbox_64bit.isChecked():
                        start_menu_shortcut_64bit = WindowsShortcut(f"{start_menu}\\Caffeine.lnk")
                        start_menu_shortcut_64bit.set_target(caffeine64_installed_location)
                        start_menu_shortcut_64bit.set_working_directory(caffeine64_install_path)
                        start_menu_shortcut_64bit.set_icon(caffeine64_copied_icon)
                        start_menu_shortcut_64bit.set_description("Caffeine 64-bit")
                        start_menu_shortcut_64bit.save()
                        
                    self.progress_bar.setValue(75)

                if self.checkbox_desktop.isChecked():
                    if self.checkbox_32bit.isChecked():
                        desktop_shortcut_32bit = WindowsShortcut(f"{public_desktop}\\Caffeine (32-bit).lnk")
                        desktop_shortcut_32bit.set_target(caffeine32_installed_location)
                        desktop_shortcut_32bit.set_working_directory(caffeine32_install_path)
                        desktop_shortcut_32bit.set_icon(caffeine32_copied_icon)
                        desktop_shortcut_32bit.set_description("Caffeine 32-bit")
                        desktop_shortcut_32bit.save()

                    if self.checkbox_64bit.isChecked():
                        desktop_shortcut_64bit = WindowsShortcut(f"{public_desktop}\\Caffeine.lnk")
                        desktop_shortcut_64bit.set_target(caffeine64_installed_location)
                        desktop_shortcut_64bit.set_working_directory(caffeine64_install_path)
                        desktop_shortcut_64bit.set_icon(caffeine64_copied_icon)
                        desktop_shortcut_64bit.set_description("Caffeine 64-bit")
                        desktop_shortcut_64bit.save()
                        
                    self.progress_bar.setValue(90)

                self.progress_bar.setValue(100)
                successfully_installed = QMessageBox.information(self, "Success", f"Successfully installed Caffeine!")
                if successfully_installed == QMessageBox.StandardButton.Ok:
                    app.quit()
            except Exception as e:
                self.progress_bar.hide()
                QMessageBox.critical(self, "Error", f"Error: Failed to install Caffeine: {str(e)}")
            finally:
                # Clean up temp files (folder where Caffeine was downloaded to)
                remove_folder(caffeine_zip_save_folder)
    
    def start_uninstall(self):
        thread = threading.Thread(target=self.uninstall, daemon=True)
        thread.start()
    
    def uninstall(self):
        self.disable_buttons()
        
        self.progress_bar.show()
        self.progress_bar.setValue(0)
        
        # Terminate Caffeine 32 and 64-bit if they're running
        caffeine_32 = ProcessManager("caffeine32.exe")
        caffeine_64 = ProcessManager("caffeine64.exe")
        
        for pid in caffeine_32.GetPid():
            if caffeine_32.Running(pid):
                caffeine_32.Kill(pid)
        self.progress_bar.setValue(5)
        
        for pid in caffeine_64.GetPid():
            if caffeine_64.Running(pid):
                caffeine_64.Kill(pid)
        self.progress_bar.setValue(10)
        
        time.sleep(3)
        
        try:
            # Remove caffeine32.exe's parent folder "Zhorn Software" in Program Files (x86)
            remove_folder(caffeine32_install_path)
            self.progress_bar.setValue(25)
            
            # Remove caffeine64.exe's parent folder "Zhorn Software" in Program Files
            remove_folder(caffeine64_install_path)
            self.progress_bar.setValue(50)
            
            # Remove shortcuts from desktop
            remove_file(os.path.join(public_desktop, "Caffeine.lnk"))
            remove_file(os.path.join(public_desktop, "Caffeine (32-bit).lnk"))
            self.progress_bar.setValue(75)
            
            # Remove shortcuts from start menu
            remove_file(os.path.join(start_menu, "Caffeine.lnk"))
            remove_file(os.path.join(start_menu, "Caffeine (32-bit).lnk"))
            
            self.progress_bar.setValue(100)
            successfully_uninstalled = QMessageBox.information(self, "Success", f"Successfully uninstalled Caffeine.")
            if successfully_uninstalled == QMessageBox.StandardButton.Ok:
                # app.quit()
                QTimer.singleShot(0, app.quit())
        except Exception as e:
            self.progress_bar.hide()
            QMessageBox.critical(self, "Error", f"Error: Failed to uninstall Caffeine: {str(e)}")
        
# Show the PyQt6 app
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = InstallerApp()
    window.show()
    sys.exit(app.exec())
