"""Autorunner."""

import subprocess  # Дождаться окончания процесса
import platform  # Определить разрядность ОС
import ctypes as ct  # Свернуть консоль
import time
import sys
import os
import os.path
import win32api  # Количество мониторов
import winshell  # Создание ярлыка
from win32com.client import Dispatch
from idlelib.tooltip import Hovertip  # Всплывающие подсказки
from mpc_hc_ini import mpcini  # Конфиг MPC
from PIL import Image, ImageTk  # Иконки на кнопках
import requests  # Скачать обновления, инсталер
import tkinter as tk
from tkinter.filedialog import askopenfilename, askopenfilenames

FILE_VERSION = '2.0.0'


def error_log(error_message):
    """Errors log."""
    work_dir = f'{os.getenv("appdata")}\\..\\local\\var\\files'
    with open(f'{work_dir}\\error.log', 'a') as err_file:
        err_file.write(time.ctime(time.time()) + 'ERROR MESSAGE: "' +
                       error_message + '"\n')


def check_update():
    """Check update."""
    url = 'https://github.com/Falc0FF/Autorunner/blob/master/version.txt'
    file = os.path.join(os.path.dirname(__file__), url.split('/')[-1])
    try:
        response = requests.get(url=url)
        with open(file, 'wb') as f:
            f.write(response.content)
    except Exception as err:
        error_log(err)
        return None
    with open(file, 'r') as f:
        new_version = f.readline()
    try:
        if new_version != FILE_VERSION:
            return new_version
    except Exception as err:
        error_log(err)
        return None
    return False


def get_update(version=None):
    """Get update."""
    if not version:
        error_log(f'new_version={version} cur_version={FILE_VERSION}')
        return
    url = 'https://github.com/Falc0FF/Autorunner/releases/download/' \
          f'v.{version}/ar.exe'
    work_dir = f'{os.getenv("appdata")}\\..\\local\\var\\files'
    newfile = os.path.join(work_dir, url.split('/')[-1])
    file = newfile.replace('\\files', '')
    try:
        response = requests.get(url=url)
        with open(newfile, 'wb') as f:
            f.write(response.content)
    except Exception as err:
        error_log(err)
        return None
    with open('update.bat', mode='w', encoding='utf-8') as upd:
        upd.write(f'''@echo off
chcp 65001>nul
echo Update complete
timeout /t 3 /nobreak>nul
move /Y "{newfile}" "{file}"
start "{file}"
del update.bat 2>nul''')
    os.startfile('update.bat', "runas")
    sys.exit()


class Application(tk.Tk):
    """App."""

    def __init__(self):
        """Create form."""
        super().__init__()
        self.form_width = 404  # Ширина окна
        self.form_height = 260  # Высота окна
        self.geometry(f'{self.form_width}x{self.form_height}')
        self.attributes('-alpha', 1)  # Прозрачность формы (0..1)
        self.attributes('-topmost', False)  # Поверх всех окон
        self.resizable(False, False)  # Изменение размеров окна
        self.title(f'Autorunner v.{FILE_VERSION}')  # Название окна
        self.set_vars()  # Создание переменных
        self.set_ui()  # Наполнение виджетами

    def set_vars(self):
        """Create variables."""
        # ---------------------Поиск установленного MPC-----------------------
        self.mpc_file = {}
        self.mpc_file[2] = 'not found'
        if platform.architecture()[0] == '32bit':
            self.mpc_file[32] = r'C:\Program Files\MPC-HC\mpc-hc.exe'
            if os.path.isfile(self.mpc_file[32]):
                self.mpc_file[1] = self.mpc_file[32]
                self.mpc_file[2] = 'found'
        else:
            self.mpc_file[64] = r'C:\Program Files\MPC-HC\mpc-hc64.exe'
            self.mpc_file[3264] = r'C:\Program Files (x86)\MPC-HC\mpc-hc.exe'
            if os.path.isfile(self.mpc_file[64]):
                self.mpc_file[1] = self.mpc_file[64]
                self.mpc_file[2] = 'found'
            elif os.path.isfile(self.mpc_file[3264]):
                self.mpc_file[1] = self.mpc_file[3264]
                self.mpc_file[2] = 'found'
        # --------------------------------------------------------------------
        self.work_dir = f'{os.getenv("appdata")}\\..\\local\\var\\files'
        self.startup_folder = os.getenv(
            'appdata') + r'\microsoft\windows\start menu\programs\startup'
        self.image_monitor = ImageTk.PhotoImage(
            Image.open(f'{os.path.dirname(__file__)}\\imgs\\monitor.ico'))
        image_pin = Image.open(f'{os.path.dirname(__file__)}\\imgs\\pin.png')
        self.image_pin = ImageTk.PhotoImage(
            image_pin.resize((20, 20), Image.Resampling.LANCZOS))
        image_unpin = Image.open(
            f'{os.path.dirname(__file__)}\\imgs\\unpin.png')
        self.image_unpin = ImageTk.PhotoImage(
            image_unpin.resize((20, 20), Image.Resampling.LANCZOS))
        self.monitor_count = len(win32api.EnumDisplayMonitors())
        self.filespath = []
        self.select_file_label = []
        self.radio_frame = []
        self.monitor_num_label = []
        self.monitor_number = []
        self.monitor_list = []
        self.run_commands_list = []
        self.files_in_monitor = {}
        for i in range(8):
            self.monitor_list.append([])
            self.files_in_monitor[i+1] = []

    def set_ui(self):
        """Create widgets."""
        self.set_ui_status()
        self.set_ui_install()
        self.set_ui_files()
        self.set_ui_monitor()
        self.set_ui_filesbut()
        self.set_ui_result()
        self.set_ui_startup()
        self.set_ui_exit()
        self.set_ui_openfolder()
        self.set_ui_clear()

    def set_ui_status(self):
        """Status label+Pin+Monitor Buttons."""
        self.mpc_text = 'MPC is not installed'
        self.mpc_text_color = 'red'
        if self.mpc_file[2] == 'found':
            self.mpc_text = 'MPC is installed'
            self.mpc_text_color = 'green'
        self.status_label = tk.Label(self, relief=tk.GROOVE,
                                     text=self.mpc_text,
                                     foreground=self.mpc_text_color)
        self.pin_label = tk.Label(self,
                                  image=self.image_pin,
                                  relief=tk.GROOVE)
        self.pin_label.bind('<Button-1>', self.app_pin)
        self.select_monitor_label = tk.Label(self,
                                             image=self.image_monitor,
                                             relief=tk.GROOVE)
        self.pin_label.place(x=0, y=0,
                             width=24, height=24)
        self.status_label.place(x=24, y=0,
                                width=self.form_width-120, height=24)
        self.select_monitor_label.place(x=self.form_width-100, y=0,
                                        width=100, height=24)
        self.mpc_tooltip = Hovertip(self.status_label,
                                    'Media Player Classic не установлен',
                                    hover_delay=100)
        if self.status_label['text'] == 'MPC is installed':
            self.mpc_tooltip.text = 'Media Player Classic установлен'
        Hovertip(self.select_monitor_label, 'Выберите номер монитора',
                 hover_delay=100)
        Hovertip(self.pin_label, 'Поверх всех окон',
                 hover_delay=100)

    def set_ui_install(self):
        """Install MPC+CFG Buttons."""
        self.install_button = tk.Button(self,
                                        text='Install MPC',
                                        command=self.app_installmpc)
        self.cfg_button = tk.Button(self,
                                    text='CFG',
                                    command=self.app_mpc_cfg)
        self.install_button.place(x=0, y=24,
                                  width=68, height=24)
        self.cfg_button.place(x=68, y=24,
                              width=28, height=24)
        self.cfg_button['state'] = 'disabled'
        if self.mpc_file[2] == 'found':
            self.install_button['state'] = 'disabled'
            self.cfg_button['state'] = 'normal'
        Hovertip(self.install_button, 'Установить Media Player Classic',
                 hover_delay=100)
        Hovertip(self.cfg_button, 'Скопировать файл настроек в папку с MPC',
                 hover_delay=100)

    def set_ui_files(self, files_list=None):
        """Select files Label."""
        if files_list:
            for i in range(len(files_list)):
                self.select_file_label.append(tk.Label(self, anchor='nw',
                                                       font='Courier 8 bold'))
                Hovertip(self.select_file_label[i], files_list[i],
                         hover_delay=100)
                file_text = files_list[i]
                if len(files_list[i]) > 23:
                    file_text = f'{files_list[i][:3]}...' \
                                f'{files_list[i][len(files_list[i])-23:]}\n'
                self.select_file_label[i]['text'] += '\n' + file_text.upper()
                self.select_file_label[i].place(x=96, y=24+i*28, width=209,
                                                height=28)
                self.radio_frame[i].place(x=self.form_width-100, y=24+i*29,
                                          width=100, height=29)

    def set_ui_monitor(self):
        """Select monitor entrys."""
        for i in range(8):
            self.radio_frame.append(tk.Frame(self))
            self.monitor_num_label.append(
                tk.Label(self.radio_frame[i], text=' 1  2  3  4  5  6  7  8',
                         bd=1, anchor='n'))
            self.monitor_num_label[i].place(x=0, y=0,
                                            width=100, height=14)
            self.monitor_number.append(tk.IntVar())
            self.monitor_number[i].set(0)
            for j in range(8):
                self.monitor_list[i].append(
                    tk.Radiobutton(self.radio_frame[i],
                                   variable=self.monitor_number[i],
                                   value=j, borderwidth=0))
                self.monitor_list[i][j].place(x=12*j, y=14)
                if j + 1 > self.monitor_count:
                    self.monitor_list[i][j]['state'] = 'disabled'

    def set_ui_filesbut(self):
        """Select files Button."""
        self.select_file_button = tk.Button(self,
                                            text='Select file',
                                            command=self.app_selectfile)
        self.select_file_button.place(x=0, y=24*2, width=96, height=36)
        self.select_file_button['state'] = 'disabled'
        if self.mpc_file[2] == 'found':
            self.select_file_button['state'] = 'normal'
        Hovertip(self.select_file_button, 'Выбрать видео файл или картинку',
                 hover_delay=100)

    def set_ui_result(self):
        """Check result Button."""
        self.check_button = tk.Button(self, text='Check result',
                                      command=self.app_check)
        self.check_button.place(x=0, y=24*2+36, width=96, height=32)
        self.check_button['state'] = 'disabled'
        Hovertip(self.check_button, 'Вывести на экран', hover_delay=100)

    def set_ui_startup(self):
        """Add shortcut to startup Button."""
        self.to_startup_button = tk.Button(
            self, text='Add shortcut\nto startup',
            command=self.app_startup)
        self.to_startup_button.place(x=0, y=24*2+36+32, width=96, height=36)
        self.to_startup_button['state'] = 'disabled'
        Hovertip(self.to_startup_button, 'Добавить ярлык в автозагрузку',
                 hover_delay=100)

    def set_ui_exit(self):
        """Run and exit Button."""
        self.exit_button = tk.Button(self, text='Run and exit',
                                     command=self.app_runexit)
        self.exit_button.place(x=0, y=self.form_height-36*3,
                               width=96, height=36)
        self.exit_button['state'] = 'disabled'
        Hovertip(self.exit_button,
                 'Запустить все ярлыки из автозагрузки и выйти',
                 hover_delay=100)

    def set_ui_openfolder(self):
        """Open startup folder."""
        self.startup_folder_button = tk.Button(self,
                                               text='Open startup\nfolder',
                                               command=self.app_startup_folder)
        self.startup_folder_button.place(x=0, y=self.form_height-36*2,
                                         width=96, height=36)
        Hovertip(self.startup_folder_button, 'Открыть папку автозагрузки',
                 hover_delay=100)

    def set_ui_clear(self):
        """Clear old shortcuts from startup."""
        self.clear_button = tk.Button(self,
                                      text='Clear old\nshortcuts',
                                      command=self.app_clear)
        self.clear_button.place(x=0, y=self.form_height-36,
                                width=96, height=36)
        Hovertip(self.clear_button, 'Очистить ранее созданные ярлыки',
                 hover_delay=100)

    def app_pin(self, event):
        """PIN/UNPIN form."""
        current_top_most = self.attributes('-topmost')
        if current_top_most == 0:
            self.attributes('-topmost', True)
            self.pin_label.config(image=self.image_unpin)
        elif current_top_most == 1:
            self.attributes('-topmost', False)
            self.pin_label.config(image=self.image_pin)

    def find_mpc_installer(self, addr=None):
        """Find MPC installer."""
        if addr:
            if os.path.isfile(addr):
                return addr
        elif not addr:
            downloads_dir = os.getenv('userprofile') + r'\downloads'
            installer_name = 'MPC-HC.1.7.9.x86.exe'
            result = os.path.join(self.work_dir, installer_name)
            if os.path.isfile(result):
                return result
            else:
                result = os.path.join(downloads_dir, installer_name)
                if os.path.isfile(result):
                    return result

    def download_mpc_installer(self):
        """Download MPC installer."""
        url = 'https://sourceforge.net/projects/mpc-hc/files/MPC%20Home' + \
            'Cinema%20-%20Win32/MPC-HC_v1.7.9_x86/MPC-HC.1.7.9.x86.exe'
        file = os.path.join(self.work_dir, url.split('/')[-1])
        try:
            response = requests.get(url=url)
            with open(file, 'wb') as f:
                f.write(response.content)
        except Exception as err:
            print(err)
        return file

    def app_installmpc(self):
        """Install MPC."""
        self.status_label['text'] = 'Waiting...'
        self.status_label.update_idletasks()
        self.mpc_installer = self.find_mpc_installer()
        if not self.mpc_installer:
            self.mpc_installer = self.find_mpc_installer(
                self.download_mpc_installer())
        while not self.mpc_installer:
            self.status_label['text'] = 'Press button on window'
            asktitle = 'MPC Installer file not found'
            askmessage = 'Инсталлер Media Player Classic не найден. ' + \
                'Укажите путь к этому файлу.'
            msgbox = tk.messagebox.askokcancel(title=asktitle,
                                               message=askmessage)
            if not msgbox:
                self.status_label['text'] = 'MPC is not installed'
                return
            else:
                self.status_label['text'] = 'Select MPC installer'
                self.mpc_installer = self.find_mpc_installer(
                    askopenfilename(filetypes=[
                        ("Applications", ".exe"), ("All types", ".*")]))
        self.status_label['text'] = 'Waiting...'
        with open('silentmpc.bat', mode='w', encoding='utf-8') as silent:
            silent.write(f'''@echo off
chcp 65001>nul
"{self.mpc_installer}" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-
del silentmpc.bat 2>nul''')
        installation = subprocess.Popen('silentmpc.bat')
        installation.wait()
        if platform.architecture()[0] == '32bit':
            self.mpc_file[1] = self.mpc_file[32]
        else:
            self.mpc_file[1] = self.mpc_file[3264]
        while not os.path.isfile(self.mpc_file[1]):
            self.status_label['text'] = 'Press button on window'
            asktitle = 'MPC file not found'
            askmessage = 'Программа Media Player Classic не найдена. ' + \
                'Укажите путь к файлу mpc-hc.exe.'
            msgbox = tk.messagebox.askokcancel(title=asktitle,
                                               message=askmessage)
            if not msgbox:
                self.status_label['text'] = 'MPC is not installed'
                return
            else:
                self.status_label['text'] = 'Select mpc-hc.exe'
                self.mpc_file[1] = askopenfilename(filetypes=[
                    ("Media Player Classic", "mpc-hc.exe;mpc-hc64.exe")])
        self.status_label['foreground'] = 'green'
        self.status_label['text'] = 'MPC is installed'
        self.mpc_tooltip.text = 'Media Player Classic установлен'
        self.cfg_button['state'] = 'normal'
        self.install_button['state'] = 'disabled'

    def create_mpc_cfg(self):
        """Create MPC CFG."""
        cfgfile = os.path.join(self.work_dir, 'mpc-hc.ini')
        with open(cfgfile, 'w') as mpccfg:
            mpccfg.write(mpcini)
        return cfgfile

    def app_mpc_cfg(self):
        """Copy MPC CFG."""
        cfgfile = self.create_mpc_cfg()
        with open(f'{cfgfile[:-10]}cfg_copy.bat', 'w',
                  encoding='utf-8') as cfgcopy:
            cfgcopy.write(f'''@echo off
chcp 65001>nul
move "{cfgfile}" "{self.mpc_file[1][:-10]}{cfgfile[-10:]}"
del {cfgfile[:-10]}cfg_copy.bat 2>nul''')
        os.startfile(f'{cfgfile[:-10]}cfg_copy.bat', "runas")
        time.sleep(4)
        if not os.path.isfile(f'{self.mpc_file[1][:-10]}{cfgfile[-10:]}'):
            asktitle = 'CFG file not found'
            askmessage = 'Не удалось переместить ini файл из текущей \
                директории в директорию с MPC. Попробуйте вручную.'
            tk.messagebox.showwarning(title=asktitle, message=askmessage)
            os.startfile(self.mpc_file[1][:-10])
        self.cfg_button['state'] = 'disabled'
        self.select_file_button['state'] = 'normal'

    def app_selectfile(self):
        """Select file."""
        filespath_old = []
        if len(self.filespath) != 0:
            del filespath_old[:]
            filespath_old = self.filespath[:]
        self.filespath = askopenfilenames()
        if len(self.filespath) != 0:
            for i in range(len(filespath_old)):
                self.select_file_label[i].place_forget()
                self.radio_frame[i].place_forget()
            del self.select_file_label[:]
            self.set_ui_files(self.filespath)
            self.check_button['state'] = 'normal'
            self.to_startup_button['state'] = 'normal'
        elif len(self.filespath) == 0:
            del self.filespath[:]
            self.filespath = filespath_old[:]

    def app_startup_folder(self):
        """Open startup folder."""
        os.startfile(self.startup_folder)

    def app_clear(self):
        """Clear old shortcuts."""
        mess = 'Вы действительно хотите удалить ранее созданные ярлыки?'
        msg = tk.messagebox.askyesno(title='Delete all shortcuts',
                                     message=mess)
        if msg:
            for file in os.listdir(self.startup_folder):
                if len(file) == 12 and file[:-4].isdigit():
                    os.remove(os.path.join(self.startup_folder, file))

    def run_command(self):
        """Run command."""
        # Очищаем словарь монитор-файлы
        for i in self.files_in_monitor.values():
            del i[:]
        # Заполняем словарь монитор-файлы
        for i in range(len(self.select_file_label)):
            monitor_num = self.monitor_number[i].get() + 1
            self.files_in_monitor[monitor_num].append(
                f'"{self.filespath[i]}"')
        del self.run_commands_list[:]
        for monitor, file in self.files_in_monitor.items():
            if file:
                files_param = ' '.join(file)
                self.run_commands_list.append(
                    f'"{self.mpc_file[1]}" {files_param} '
                    f'/new /play /fullscreen /monitor {monitor}'
                )
        return self.run_commands_list

    def app_check(self):
        """Check result."""
        for command in self.run_command():
            subprocess.Popen(command)

    def app_desktop(self, cmd):
        """Make shortcut on desktop."""
        shortcut_name = str(round(time.time()*100000))[7:]
        desktop = winshell.desktop()
        path = os.path.join(desktop, f"{shortcut_name}.lnk")
        target = f"{self.mpc_file[1]} "
        wDir = f"{os.path.dirname(self.mpc_file[1])}"
        icon = f"{self.mpc_file[1]}"
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = target
        shortcut.Arguments = cmd.split('.exe" ')[-1]
        shortcut.WorkingDirectory = wDir
        shortcut.IconLocation = icon
        shortcut.save()
        self.exit_button['state'] = 'normal'
        return path

    def app_startup(self):
        """Add shortcut to startup."""
        with open(f'{winshell.desktop()}\\file_move.bat', 'w',
                  encoding='utf-8') as filemove:
            filemove.write('''@echo off
chcp 65001>nul''')
            for command in self.run_command():
                shortcutfile = self.app_desktop(command)
                file_name = shortcutfile.split('\\')[-1]
                filemove.write(f'''
move "{shortcutfile}" "{self.startup_folder}\\{file_name}">nul 2>nul''')
            filemove.write(f'''
del {os.path.dirname(shortcutfile)}\\file_move.bat 2>nul''')
        os.startfile(f'{os.path.dirname(shortcutfile)}\\file_move.bat')

    def app_runexit(self):
        """Run and Exit."""
        for file in os.listdir(self.startup_folder):
            if len(file) == 12 and file[:-4].isdigit():
                os.startfile(os.path.join(self.startup_folder, file))
        self.destroy()
        sys.exit()


def main():
    """Basic_function."""
    root = Application()
    root.iconify()
    root.update()
    root.deiconify()
    root.mainloop()


if __name__ == '__main__':
    ct.windll.user32.ShowWindow(ct.windll.kernel32.GetConsoleWindow(), 6)
    if len(sys.argv) == 2 and '-ver' in sys.argv:
        print(FILE_VERSION)
        with open('file_version.txt', 'w') as fver:
            fver.write(FILE_VERSION)
    elif len(sys.argv) == 2 and '-upd' in sys.argv:
        get_update(check_update())
    elif len(sys.argv) == 2 and '-test' in sys.argv:
        print(winshell.desktop())
    elif len(sys.argv) < 2:
        main()
