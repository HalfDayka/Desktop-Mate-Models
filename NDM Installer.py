import os
import threading
import subprocess
import winreg
from tkinter import filedialog
import customtkinter as ctk
import requests
import zipfile
import shutil
import winshell
import locale
import pyperclip
from win32com.client import Dispatch
import pkg_resources
from bs4 import BeautifulSoup

def find_steam_path():
    try:
        reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\\WOW6432Node\\Valve\\Steam")
        steam_path, _ = winreg.QueryValueEx(reg_key, "InstallPath")
        if os.path.exists(os.path.join(steam_path, "steam.exe")):
            return steam_path
    except FileNotFoundError:
        pass

    possible_paths = [
        os.path.join(os.getenv("ProgramFiles(x86)"), "Steam"),
        os.path.join(os.getenv("ProgramFiles"), "Steam"),
        os.path.join(os.getenv("LocalAppData"), "Steam"),
        os.path.join(os.getenv("UserProfile"), "AppData", "Roaming", "Steam"),
        "C:\\Steam"
    ]

    for drive in 'ABCDEFGHIJKLMNOPQRSTUVWXYZ':
        possible_paths.append(f"{drive}:\\Program Files (x86)\\Steam")
        possible_paths.append(f"{drive}:\\Steam")

    for path in possible_paths:
        if os.path.exists(os.path.join(path, "steam.exe")):
            return path
    return None

def execute_threaded(func):
    thread = threading.Thread(target=func)
    thread.start()

import webbrowser

class DesktopMateInstallerApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Nagatoro Desktop Mate Installer")
        self.geometry("800x400")
        self.resizable(False, False)

        self.steam_path = find_steam_path()
        self.current_step = 0
        self.language = self.get_system_language()
        self.translations = {
            'ru': {
                "info_label": "Информационный раздел",
                "button_labels": [
                    "1. Установить Desktop Mate", "2. Установить .NET Runtime", "3. Скачать MelonLoader",
                    "4. Скачать Custom Avatar Loader", "5. Установить MelonLoader", "6. Настроить параметры запуска",
                    "7. Распаковать модели Нагаторо", "8. Добавить в автозапуск", "9. Запустить Desktop Mate"
                ],
                "steps": [
                    "Проверим установлен ли Desktop Mate, если нет, то начнется загрузка.\nНажмите на зеленую кнопку для продолжения.",
                    "Установим .NET Runtime для корректной работы приложения.",
                    "Скачаем MelonLoader.",
                    "Скачаем и установим Custom Avatar Loader.",
                    "Установим MelonLoader в папку Desktop Mate.",
                    "Настройте параметры запуска Desktop Mate.",
                    "Распакуем модели Нагаторо.",
                    "Добавим автозапуск программы.",
                    "Запустим Desktop Mate."
                ]
            },
            'en': {
                "info_label": "Information Section",
                "button_labels": [
                    "1. Install Desktop Mate", "2. Install .NET Runtime", "3. Download MelonLoader",
                    "4. Download Custom Avatar Loader", "5. Install MelonLoader", "6. Set Launch Parameters",
                    "7. Unpack Nagatoro Models", "8. Add to autorun", "9. Launch Desktop Mate"
                ],
                "steps": [
                    "Let's check if Desktop Mate is installed, if not, it will start booting.\nClick the green button to continue.",
                    "Let's install the .NET Runtime for the application to run correctly.",
                    "Download MelonLoader.",
                    "Download and install Custom Avatar Loader.",
                    "Install MelonLoader in the Desktop Mate folder.",
                    "Set launch parameters for Desktop Mate.",
                    "Unpack Nagatoro models.",
                    "Let's add an autorun program.",
                    "Launch Desktop Mate."
                ]
            }
        }

        self.current_translation = self.translations.get(self.language, self.translations['en'])

        self.left_frame = ctk.CTkFrame(self, width=200, corner_radius=10)
        self.left_frame.pack(side="left", fill="y", padx=10, pady=10)

        self.right_frame = ctk.CTkFrame(self, corner_radius=10)
        self.right_frame.pack(side="right", fill="both", expand=True, padx=10, pady=10)

        self.info_label = ctk.CTkLabel(self.right_frame, text=self.current_translation["info_label"], wraplength=400, justify="left")
        self.info_label.pack(padx=10, pady=10, fill="both", expand=True)

        self.create_buttons()
        self.update_advisor()
        self.add_author_links_to_buttons()

    def add_author_links_to_buttons(self):
        bottom_links_frame = ctk.CTkFrame(self.left_frame, fg_color="transparent")
        bottom_links_frame.pack(side="bottom", fill="x", pady=5)

        bottom_links_frame.grid_rowconfigure(0, weight=1)
        bottom_links_frame.grid_columnconfigure(0, weight=1)
        bottom_links_frame.grid_columnconfigure(1, weight=1)
        bottom_links_frame.grid_columnconfigure(2, weight=1)
        bottom_links_frame.grid_columnconfigure(3, weight=1)
        bottom_links_frame.grid_columnconfigure(4, weight=1)

        vk_url = "https://vk.com/ijiranaidenagatorosan"
        telegram_url = "https://t.me/Nagatoro"

        halfday_label = ctk.CTkLabel(bottom_links_frame, text="Nagatoro:", fg_color="transparent")
        halfday_label.grid(row=0, column=1, padx=3)

        vk_label = ctk.CTkLabel(bottom_links_frame, text="VK", cursor="hand2")
        vk_label.grid(row=0, column=2, padx=3)
        vk_label.bind("<Button-1>", lambda e: webbrowser.open(vk_url))

        telegram_label = ctk.CTkLabel(bottom_links_frame, text="Telegram", cursor="hand2")
        telegram_label.grid(row=0, column=3, padx=3)
        telegram_label.bind("<Button-1>", lambda e: webbrowser.open(telegram_url))

        for widget in bottom_links_frame.winfo_children():
            widget.grid_configure(sticky="nsew")

    def get_system_language(self):
        language, encoding = locale.getdefaultlocale()
        if not language:
            return 'en'
        if 'ru' in language:
            return 'ru'
        else:
            return 'en'

    def toggle_language(self):
        # Переключаем язык
        self.language = 'en' if self.language == 'ru' else 'ru'
        self.current_translation = self.translations.get(self.language, self.translations['en'])

        self.info_label.configure(text=self.current_translation["info_label"])
        for idx, btn in enumerate(self.button_widgets):
            btn.configure(text=self.current_translation["button_labels"][idx])
        self.update_advisor()

    def create_buttons(self):
        self.buttons = [
            ("1. Установить Desktop Mate", self.handle_install_desktop_mate),
            ("2. Установить .NET Runtime", self.handle_install_dotnet),
            ("3. Скачать MelonLoader", self.handle_install_melonloader),
            ("4. Скачать Custom Avatar Loader", self.handle_install_custom_avatar),
            ("5. Установить MelonLoader", self.handle_setup_melonloader),
            ("6. Настроить параметры запуска", self.handle_set_launch_params),
            ("7. Распаковать модели Нагаторо", self.handle_install_nagatoro_models),
            ("8. Добавить в автозапуск", self.add_to_autorun),
            ("9. Запустить Desktop Mate", self.handle_launch_desktop_mate)
        ]

        self.button_widgets = []
        for idx, (text, command) in enumerate(self.buttons):
            btn = ctk.CTkButton(self.left_frame, text=self.current_translation["button_labels"][idx], command=lambda cmd=command, step=idx: self.execute_and_update(cmd, step))
            btn.pack(fill="x", padx=5, pady=5)
            self.button_widgets.append(btn)

        self.language_button = ctk.CTkButton(self.right_frame, text="Change Language", command=self.toggle_language)
        self.language_button.pack(padx=10, pady=10)

    def update_advisor(self):
        for idx, btn in enumerate(self.button_widgets):
            if idx == self.current_step:
                btn.configure(fg_color="green")
            else:
                btn.configure(fg_color="transparent")

        self.info_label.configure(text=self.current_translation["steps"][self.current_step])

    def execute_and_update(self, command, step):
        execute_threaded(command)

        if step == self.current_step:
            self.current_step += 1

    def update_info(self, message):
        self.info_label.configure(
            text=message,
            # justify="center",
            font=("Helvetica", 14)
        )

    def handle_install_desktop_mate(self):
        if self.steam_path:
            if self.language == "ru":
                self.update_info("Проверяем, установлена ли Desktop Mate...")
            else:
                self.update_info("Checking if Desktop Mate is installed...")

            app_id = "3301060"
            acf_path = os.path.join(self.steam_path, "steamapps", f"appmanifest_{app_id}.acf")

            if os.path.exists(acf_path):
                if self.language == "ru":
                    self.update_info("Desktop Mate уже установлен.")
                else:
                    self.update_info("Desktop Mate is already installed.")

                if self.current_step == 1:
                    self.update_advisor()
                return
            else:
                if self.language == "ru":
                    self.update_info("Устанавливаем Desktop Mate...")
                else:
                    self.update_info("Installing Desktop Mate...")
                subprocess.run([os.path.join(self.steam_path, "steam.exe"), f"steam://install/{app_id}"])
        else:
            if self.language == "ru":
                self.update_info("Не удалось найти папку Steam. Пожалуйста, проверьте установку Steam.")
            else:
                self.update_info("Failed to find Steam folder. Please check your Steam installation.")

    def handle_install_dotnet(self):
        try:
            if not hasattr(self, 'step_1_done'):
                dotnet_installed = subprocess.run(["dotnet", "--version"], capture_output=True).returncode == 0
                if dotnet_installed:
                    if self.current_step == 2:
                        self.update_advisor()
                    return

                self.step_1_done = True

                if self.language == "ru":
                    self.update_info("Скачивание .NET 6.0 Desktop Runtime (v6.0.36)...")
                else:
                    self.update_info("Download .NET 6.0 Desktop Runtime (v6.0.36)...")

                installer_path = "windowsdesktop-runtime-6.0.36-win-x64.exe"

                # Если файл существует, запускаем установку
                if os.path.exists(installer_path):
                    subprocess.run([installer_path], check=True)
                else:
                    # Если файл не найден, скачиваем его
                    self.download_and_install_dotnet()

            else:
                del self.step_1_done
                if self.current_step == 2:
                    self.update_advisor()

        except Exception:
            pass

    def download_and_install_dotnet(self):
        try:
            # URL страницы, с которой мы будем парсить ссылку
            url = 'https://dotnet.microsoft.com/en-us/download/dotnet/thank-you/runtime-desktop-6.0.36-windows-x64-installer'

            # Получаем HTML страницы
            response = requests.get(url)
            if response.status_code == 200:
                # Парсим страницу
                soup = BeautifulSoup(response.text, 'html.parser')

                # Ищем ссылку для скачивания
                download_link = soup.find('a', {'id': 'directLink'})['href']
                if download_link:
                    print(f"Ссылка для скачивания: {download_link}")

                    # Скачиваем файл
                    download_response = requests.get(download_link)
                    installer_path = "windowsdesktop-runtime-6.0.36-win-x64.exe"

                    # Сохраняем файл
                    with open(installer_path, 'wb') as f:
                        f.write(download_response.content)

                    print("Файл успешно скачан.")

                    if self.language == "ru":
                        self.update_info("По окончанию установки нажмите на зеленую кнопку еще раз.")
                    else:
                        self.update_info("When the installation is complete, press the green button again.")

                    # Запускаем установку
                    subprocess.run([installer_path], check=True)
                else:
                    print("Не удалось найти ссылку для скачивания.")
            else:
                print("Не удалось загрузить страницу.")
        except Exception as e:
            print(f"Ошибка при скачивании и установке .NET: {e}")

    def handle_install_melonloader(self):
        if self.language == "ru":
            self.update_info("Скачиваем актуальную версию MelonLoader...")
        else:
            self.update_info("Download the current version of MelonLoader...")

        url = "https://github.com/LavaGang/MelonLoader/releases/latest/download/MelonLoader.Installer.exe"
        response = requests.get(url)
        installer_path = "MelonLoader.Installer.exe"
        with open(installer_path, "wb") as file:
            file.write(response.content)

        if self.language == "ru":
            self.update_info("MelonLoader успешно скачан.")
        else:
            self.update_info("MelonLoader has been successfully downloaded.")

        if self.current_step == 3:
            self.update_advisor()

    def handle_install_custom_avatar(self):
        install_path = os.path.join(self.steam_path, "steamapps", "common", "Desktop Mate")

        if self.language == "ru":
            self.update_info("Скачиваем и устанавливаем актуальную версию Custom Avatar Loader...")
        else:
            self.update_info("Download and install the current version of Custom Avatar Loader...")

        url = "https://github.com/YusufOzmen01/desktopmate-custom-avatar-loader/releases/latest/download/CustomAvatarLoader.zip"
        response = requests.get(url)
        zip_path = "CustomAvatarLoader.zip"
        with open(zip_path, "wb") as file:
            file.write(response.content)

        if self.current_step == 4:
            self.update_advisor()

        try:
            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(install_path)
            if self.language == "ru":
                self.update_info("Custom Avatar Loader успешно установлен.")
            else:
                self.update_info("Custom Avatar Loader has been successfully installed.")
        except PermissionError:
            if self.language == "ru":
                self.update_info("Ошибка доступа при установке Custom Avatar Loader. Попробуйте запустить с правами администратора.")
            else:
                self.update_info("Access error during Custom Avatar Loader installation. Try running as administrator.")

    def handle_setup_melonloader(self):
        if self.language == "ru":
            self.update_info("Инструкция для установки MelonLoader:\n\n- В списке приложений выберите Desktop Mate\n- Выберите последнюю версию MelonLoader\n- Установите\n- Закройте MelonLoader.")
        else:
            self.update_info("MelonLoader installation instructions:\n\n- In the application list, select Desktop Mate\n- Choose the latest version of MelonLoader\n- Install it\n- Close MelonLoader.")

        subprocess.run(["MelonLoader.Installer.exe"])

        if self.current_step == 5:
            self.update_advisor()

    def handle_set_launch_params(self):
        if not hasattr(self, 'step_1_done'):
            if self.language == "ru":
                instruction_text = (
                    "Для отключения консоли измените параметры запуска программы в Steam:\n"
                    "--melonloader.hideconsole (этот текст уже скопирован в буфер обмена)\n\n"
                    "Для этого:\n"
                    "- Зайдите в библиотеку Steam\n"
                    "- Нажмите правой кнопкой мыши на Desktop Mate\n"
                    "- Откройте Свойства\n"
                    "- Параметры запуска уже скопированы, просто вставьте"
                )
            else:
                instruction_text = (
                    "To disable the console, change the launch parameters of the program in Steam:\n"
                    "--melonloader.hideconsole (this text has been copied to the clipboard)\n\n"
                    "To do this:\n"
                    "- Go to your Steam library\n"
                    "- Right-click on Desktop Mate\n"
                    "- Open Properties\n"
                    "- The startup parameters are already copied, just paste"
                )

            pyperclip.copy("--melonloader.hideconsole")

            self.update_info(instruction_text)
            self.step_1_done = True
        else:
            del self.step_1_done

            if self.current_step == 6:
                self.update_advisor()

    def handle_install_nagatoro_models(self):
        try:
            if self.language == "ru":
                self.update_info("Скачиваем и распаковываем модели из репозитория GitHub...")
            else:
                self.update_info("Downloading and unpacking models from the GitHub repository...")

            # Указываем путь к папке, куда будем скачивать репозиторий
            models_dir = os.path.join(os.path.expanduser("~"), "Documents", "Models")
            vrm_dir = os.path.join(models_dir, "vrm")
            repo_url = "https://github.com/HalfDayka/NDM-Models/archive/refs/heads/main.zip"

            # Если папка Models не существует, создаем её
            if not os.path.exists(models_dir):
                os.makedirs(models_dir)

            # Если папка vrm не существует, создаем её
            if not os.path.exists(vrm_dir):
                os.makedirs(vrm_dir)

            # Скачиваем архив репозитория
            zip_path = os.path.join(models_dir, "NDM-Models.zip")
            self.update_info("Скачиваем архив репозитория...")
            response = requests.get(repo_url, stream=True)
            response.raise_for_status()  # Проверяем успешность запроса

            with open(zip_path, "wb") as zip_file:
                for chunk in response.iter_content(chunk_size=8192):
                    zip_file.write(chunk)

            # Распаковываем архив
            self.update_info("Распаковываем архив репозитория...")
            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(models_dir)

            # Перемещаем модели в нужную папку
            extracted_dir = os.path.join(models_dir, "NDM-Models-main", "vrm")
            if not os.path.exists(extracted_dir):
                raise FileNotFoundError("Не найдена папка с моделями в скачанном архиве.")

            for file in os.listdir(extracted_dir):
                if file.endswith(".vrm"):
                    shutil.move(
                        os.path.join(extracted_dir, file),
                        os.path.join(vrm_dir, file)
                    )

            # Удаляем архив и временную папку
            os.remove(zip_path)
            shutil.rmtree(os.path.join(models_dir, "NDM-Models-main"), ignore_errors=True)

            # Устанавливаем иконки для .vrm файлов
            self.update_info("Настраиваем иконки для файлов моделей...")
            icons_dir = pkg_resources.resource_filename(__name__, 'Icons')

            # Если папка с иконками не существует, пропускаем
            if not os.path.exists(icons_dir):
                return

            vrm_files = [f for f in os.listdir(vrm_dir) if f.endswith(".vrm")]

            if not vrm_files:
                if self.language == "ru":
                    self.update_info("Не найдено .vrm файлов в папке Models.")
                else:
                    self.update_info("No .vrm files found in the Models folder.")
                return

            # Обрабатываем каждый .vrm файл
            for vrm_file in vrm_files:
                vrm_path = os.path.join(vrm_dir, vrm_file)
                icon_name = os.path.splitext(vrm_file)[0] + ".ico"
                icon_path = os.path.join(icons_dir, icon_name)

                if not os.path.exists(icon_path):
                    continue

                # Создаем ярлыки для моделей
                shortcut_name = os.path.splitext(vrm_file)[0] + ".lnk"
                shortcut_path = os.path.join(models_dir, shortcut_name)

                try:
                    with winshell.shortcut(shortcut_path) as shortcut:
                        shortcut.path = vrm_path
                        shortcut.icon_location = (icon_path, 0)

                except Exception as e:
                    pass

            # Устанавливаем путь к модели в конфигурационный файл
            user_profile = os.environ["USERPROFILE"]
            vrm_path = os.path.join(user_profile, "Documents", "Models", "vrm", "Nagatoro Uniform.vrm")

            config_content = f"""[settings]
        disable_log_readonly = false
        vrmPath = '{vrm_path}'"""

            output_dir = os.path.join(self.steam_path, "steamapps", "common", "Desktop Mate", "UserData")
            os.makedirs(output_dir, exist_ok=True)  # Создаём директории, если их нет

            output_file = os.path.join(output_dir, "MelonPreferences.cfg")

            with open(output_file, "w", encoding="utf-8") as file:
                file.write(config_content)

            if self.language == "ru":
                self.update_info("Работа с моделями завершена.\nМодели Нагаторо успешно добавлены.")
            else:
                self.update_info("Model handling completed.\nNagatoro models added successfully.")

            if self.current_step == 7:
                self.update_advisor()

        except Exception as e:
            if self.language == "ru":
                self.update_info(f"Произошла ошибка: {e}")
            else:
                self.update_info(f"An error occurred: {e}")

    def add_to_autorun(self):
        try:
            desktop_mate_path = os.path.join(self.steam_path, "steamapps", "common", "Desktop Mate", "DesktopMate.exe")

            if not os.path.exists(desktop_mate_path):
                raise FileNotFoundError(f"Файл не найден: {desktop_mate_path}")

            startup_folder = os.path.join(os.getenv('APPDATA'), "Microsoft\\Windows\\Start Menu\\Programs\\Startup")

            shortcut_name = "DesktopMate.lnk"
            shortcut_path = os.path.join(startup_folder, shortcut_name)

            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortcut(shortcut_path)
            shortcut.TargetPath = desktop_mate_path
            shortcut.WorkingDirectory = os.path.dirname(desktop_mate_path)
            shortcut.IconLocation = desktop_mate_path  # Устанавливаем иконку ярлыка
            shortcut.Save()

            if self.language == "ru":
                self.update_info("Desktop Mate добавлен в автозапуск.")
            else:
                self.update_info("Desktop Mate added to startup.")
        except Exception as e:
            # Сообщение об ошибке
            if self.language == "ru":
                self.update_info(f"Ошибка при добавлении в автозапуск: {e}")
            else:
                self.update_info(f"Error adding to startup: {e}")

        if self.current_step == 8:
            self.update_advisor()

    def handle_launch_desktop_mate(self):
        try:
            if self.language == "ru":
                self.update_info("Запускаем Desktop Mate через Steam... Первый запуск может занять некоторое время.")
            else:
                self.update_info("Launching Desktop Mate through Steam... The first launch may take some time.")
            app_id = "3301060"
            subprocess.run([os.path.join(self.steam_path, "steam.exe"), f"steam://rungameid/{app_id}"])
            if self.language == "ru":
                self.update_info("Desktop Mate успешно запущен.\nПервый запуск может занять время.\n\nСмена модели:\n- Нажмите на персонажа левой кнопкой мыши (ЛКМ).\n- Нажмите F4, у вас откроется папка Documents.\n- Откройте папку Models и выберите любую модельку.\n\nЕсли вы пропустили, то для отключения консоли выполните пункт 6.")
            else:
                self.update_info("Desktop Mate successfully launched.\nThe first launch can take time.\n\nChange model:\n- Left-click on the character.\n- Press F4 to open the Documents folder.\n- Open the Models folder and choose any model.\n\nIf you missed it, follow step 6 to disable the console.")
        except Exception as e:
            if self.language == "ru":
                self.update_info(f"Ошибка при запуске Desktop Mate: {e}")
            else:
                self.update_info(f"Error launching Desktop Mate: {e}")

if __name__ == "__main__":
    app = DesktopMateInstallerApp()
    app.mainloop()
