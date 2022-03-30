"""Autorunner."""
# ver. 1.1

import subprocess
import platform
import time
import sys
import os
import os.path
import winshell
from win32com.client import Dispatch
from getpass import getpass
from idlelib.tooltip import Hovertip
from dotenv import load_dotenv
from mpc_hc_ini import mpcini
import requests
import tkinter as tk
from tkinter import ttk
from tkinter.filedialog import askopenfilename


def get_from_env(key):
    """Get a secret key from a file."""
    dotenv_path = os.path.join(os.path.dirname(__file__), '.env')
    load_dotenv(dotenv_path)
    return os.environ.get(key)


class Application(tk.Tk):
    """App."""

    def __init__(self):
        """Create form."""
        tk.Tk.__init__(self)
        self.geometry('236x172')  # Ширина х высота
        self.attributes('-alpha', 1)  # Прозрачность формы (0..1)
        self.attributes('-topmost', False)  # Поверх всех окон
        self.resizable(False, False)  # Изменение размеров окна
        self.title('Autorunner')
        self.set_ui()  # Наполнение виджетами

    def set_ui(self):
        """Create widgets."""
        self.set_ui_install()
        self.set_ui_file()
        self.set_ui_monitor()
        self.set_ui_result()
        self.set_ui_shortcut()
        self.set_ui_startup()
        self.set_ui_exit()

    def set_ui_install(self):
        """Install MPC Button+Label."""
        if platform.architecture()[0] == '32bit':
            self.mpc_file = r'C:\Program Files\MPC-HC\mpc-hc.exe'
        else:
            self.mpc_file = r'C:\Program Files (x86)\MPC-HC\mpc-hc.exe'
        self.install_frame = ttk.Frame(self)
        self.install_frame.pack(fill=tk.X)
        self.install_button = ttk.Button(self.install_frame,
                                         text='Install MPC',
                                         command=self.app_installmpc)
        self.cfg_button = ttk.Button(self.install_frame,
                                     text='CFG',
                                     command=self.app_mpc_cfg)
        self.mpc_text = 'MPC is not installed'
        self.mpc_text_color = 'red'
        self.cfg_button['state'] = 'disabled'
        if os.path.isfile(self.mpc_file):
            self.mpc_text = 'MPC is installed'
            self.mpc_text_color = 'green'
            self.cfg_button['state'] = 'enabled'
            self.install_button['state'] = 'disabled'
        self.install_label = ttk.Label(self.install_frame,
                                       text=self.mpc_text,
                                       foreground=self.mpc_text_color)
        self.install_button.pack(side=tk.LEFT)
        self.install_label.pack(side=tk.LEFT)
        self.cfg_button.pack(side=tk.LEFT)
        self.mpc_tooltip = Hovertip(self.install_label,
                                    'Media Player Classic не установлен',
                                    hover_delay=100)
        if self.install_label['text'] == 'MPC is installed':
            self.mpc_tooltip.text = 'Media Player Classic установлен'
        Hovertip(self.install_button, 'Установить Media Player Classic',
                 hover_delay=100)
        Hovertip(self.cfg_button, 'Скопировать файл настроек в папку с MPC',
                 hover_delay=100)

    def set_ui_file(self):
        """Select file Button+Label."""
        self.select_file_frame = ttk.Frame(self)
        self.select_file_frame.pack(fill=tk.X)
        self.select_file_button = ttk.Button(self.select_file_frame,
                                             text='Select file',
                                             command=self.app_selectfile)
        self.select_file_button.pack(side=tk.LEFT)
        self.select_file_label = ttk.Label(self.select_file_frame,
                                           text='File location')
        self.select_file_label.pack(side=tk.LEFT)
        if self.install_label['text'] == 'MPC is not installed':
            self.select_file_button['state'] = 'disabled'
        Hovertip(self.select_file_button, 'Выбрать видео файл или картинку',
                 hover_delay=100)
        Hovertip(self.select_file_label, 'Название файла', hover_delay=100)

    def set_ui_monitor(self):
        """Select monitor Button+Label."""
        self.select_monitor_frame = ttk.Frame(self)
        self.select_monitor_frame.pack(fill=tk.X)
        self.select_monitor_label = ttk.Label(
            self.select_monitor_frame,
            text='Select monitor number (1-8):')
        self.select_monitor_label.pack(side=tk.LEFT)
        self.select_monitor_entry = ttk.Entry(self.select_monitor_frame,)
        self.select_monitor_entry.pack(side=tk.LEFT)
        self.select_monitor_entry.insert(0, '1')
        Hovertip(self.select_monitor_label, 'Введите номер монитора',
                 hover_delay=100)

    def set_ui_result(self):
        """Check result Button."""
        self.result_frame = ttk.Frame(self)
        self.result_frame.pack(fill=tk.X)
        self.check_button = ttk.Button(self.result_frame,
                                       text='      Check result        ',
                                       command=self.app_check)
        self.check_button.pack(side=tk.LEFT)
        self.check_button['state'] = 'disabled'
        self.startup_folder = ttk.Button(self.result_frame,
                                         text=' Open startup folder  ',
                                         command=self.app_startup_folder)
        self.startup_folder.pack(side=tk.RIGHT)
        Hovertip(self.check_button, 'Вывести на экран', hover_delay=100)
        Hovertip(self.startup_folder, 'Открыть папку автозагрузки',
                 hover_delay=100)

    def app_clear(self):
        """Clear old shortcuts."""
        mess = 'Вы действительно хотите удалить ранее созданные ярлыки?'
        msg = tk.messagebox.askyesno(title='Delete all shortcuts',
                                     message=mess)
        if msg:
            dirname = os.getenv('appdata') + \
                r'\microsoft\windows\start menu\programs\startup'
            for file in os.listdir(dirname):
                if len(file) == 12 and file[:-4].isdigit():
                    os.remove(os.path.join(dirname, file))

    def set_ui_shortcut(self):
        """Make shortcut on desktop Button."""
        self.shortcut_frame = ttk.Frame(self)
        self.shortcut_frame.pack(fill=tk.X)
        self.on_desktop_button = ttk.Button(
            self.shortcut_frame, text='  Make shortcut on desktop  ',
            command=self.app_desktop)
        self.on_desktop_button.pack(side=tk.LEFT)
        self.on_desktop_button['state'] = 'disabled'
        self.clear_button = ttk.Button(self.shortcut_frame, text='Clear',
                                       command=self.app_clear)
        self.clear_button.pack(side=tk.RIGHT)
        Hovertip(self.on_desktop_button, 'Создать ярлык на рабочем столе',
                 hover_delay=100)
        Hovertip(self.clear_button, 'Очистить ранее созданные ярлыки',
                 hover_delay=100)

    def set_ui_startup(self):
        """Add shortcut to startup Button."""
        self.to_startup_button = ttk.Button(
            self, text='Add shortcut to startup',
            command=self.app_startup)
        self.to_startup_button.pack(fill=tk.X)
        self.to_startup_button['state'] = 'disabled'
        Hovertip(self.to_startup_button, 'Добавить ярлык в автозагрузку',
                 hover_delay=100)

    def set_ui_exit(self):
        """Run and exit Button."""
        self.exit_button = ttk.Button(self, text='Run and exit',
                                      command=self.app_runexit)
        self.exit_button.pack(fill=tk.X)
        self.exit_button['state'] = 'disabled'
        Hovertip(self.exit_button,
                 'Запустить все ярлыки из автозагрузки и выйти',
                 hover_delay=100)

    def find_mpc_installer(addr=None):
        """Find MPC installer."""
        if addr:
            if os.path.isfile(addr):
                return addr
        elif not addr:
            current_dir = os.path.dirname(__file__)
            downloads_dir = os.getenv('userprofile') + r'\downloads'
            installer_name = 'MPC-HC.1.7.9.x86.exe'
            result = os.path.join(current_dir, installer_name)
            if os.path.isfile(result):
                return result
            else:
                result = os.path.join(downloads_dir, installer_name)
                if os.path.isfile(result):
                    return result

    def download_mpc_installer():
        """Download MPC installer."""
        url = 'https://sourceforge.net/projects/mpc-hc/files/MPC%20Home' + \
            'Cinema%20-%20Win32/MPC-HC_v1.7.9_x86/MPC-HC.1.7.9.x86.exe'
        try:
            response = requests.get(url=url)
            with open(url.split('/')[-1], 'wb') as file:
                file.write(response.content)
        except Exception as err:
            print(err)
        dir = os.path.dirname(__file__)
        return os.path.join(dir, url.split('/')[-1])

    def app_installmpc(self):
        """Install MPC."""
        self.mpc_installer = Application.find_mpc_installer()
        if not self.mpc_installer:
            self.mpc_installer = Application.find_mpc_installer(
                Application.download_mpc_installer())
        while not self.mpc_installer:
            asktitle = 'MPC Installer file not found'
            askmessage = 'Инсталлер Media Player Classic не найден. ' + \
                'Укажите путь к этому файлу.'
            msgbox = tk.messagebox.askokcancel(title=asktitle,
                                               message=askmessage)
            if not msgbox:
                return
            else:
                self.mpc_installer = Application.find_mpc_installer(
                    askopenfilename())
        with open('silentmpc.bat', mode='w', encoding='1251') as silent:
            silent.write(f'''@echo off
chcp 1251>nul
"{self.mpc_installer}" /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-
del silentmpc.bat 2>nul''')
        installation = subprocess.Popen('silentmpc.bat')
        installation.wait()
        if os.path.isfile(self.mpc_file):
            self.install_label['foreground'] = 'green'
            self.install_label['text'] = 'MPC is installed'
            self.mpc_tooltip.text = 'Media Player Classic установлен'
            self.cfg_button['state'] = 'enabled'
            self.install_button['state'] = 'disabled'

    def create_mpc_cfg(self):
        """Create MPC CFG."""
        current_dir = os.path.dirname(__file__)
        cfgfile = os.path.join(current_dir, 'mpc-hc.ini')
        with open(cfgfile, 'w') as mpccfg:
            mpccfg.write(mpcini)
        return cfgfile

    def app_mpc_cfg(self):
        """Copy MPC CFG."""
        cfgfile = self.create_mpc_cfg()
        with open(f'{cfgfile[:-10]}cfg_copy.bat', 'w',
                  encoding='1251') as cfgcopy:
            cfgcopy.write(f'''@echo off
chcp 1251>nul
move "{cfgfile}" "{self.mpc_file[:-10]}{cfgfile[-10:]}"
del {cfgfile[:-10]}cfg_copy.bat 2>nul''')
        os.startfile(f'{cfgfile[:-10]}cfg_copy.bat', "runas")
        time.sleep(4)
        if not os.path.isfile(f'{self.mpc_file[:-10]}{cfgfile[-10:]}'):
            asktitle = 'CFG file not found'
            askmessage = 'Не удалось переместить ini файл из текущей \
                директории в директорию с MPC. Попробуйте вручную.'
            tk.messagebox.showwarning(title=asktitle, message=askmessage)
            os.startfile(self.mpc_file[:-10])
        self.cfg_button['state'] = 'disabled'
        self.select_file_button['state'] = 'enabled'

    def app_selectfile(self):
        """Select file."""
        self.filepath = askopenfilename()
        self.select_file_label['text'] = self.filepath[len(self.filepath)-27:]
        if self.select_file_label['text'] == '':
            self.select_file_label['text'] = 'File location'
        if self.select_file_label['text'] != 'File location':
            self.check_button['state'] = 'enabled'
            self.on_desktop_button['state'] = 'enabled'
            self.to_startup_button['state'] = 'enabled'

    def app_startup_folder(self):
        """Open startup folder."""
        dirname = os.getenv('appdata') + \
            r'\microsoft\windows\start menu\programs\startup'
        os.startfile(dirname)

    def run_command(self):
        """Run command."""
        monitor_num = self.select_monitor_entry.get()
        if monitor_num not in list(map(lambda x: str(x), range(1, 9))):
            monitor_num = '1'
        return f'"{self.mpc_file}" "{self.filepath}" ' \
               f'/new /play /fullscreen /monitor {monitor_num}'

    def app_check(self):
        """Check result."""
        subprocess.Popen(self.run_command())

    def app_desktop(self):
        """Make shortcut on desktop."""
        self.exit_button['state'] = 'enabled'
        shortcut_name = str(round(time.time()*100000))[7:]
        desktop = winshell.desktop()
        path = os.path.join(desktop, f"{shortcut_name}.lnk")
        target = f"{self.mpc_file} "
        wDir = f"{self.mpc_file[:-10]}"
        icon = f"{self.mpc_file}"
        shell = Dispatch('WScript.Shell')
        shortcut = shell.CreateShortCut(path)
        shortcut.Targetpath = target
        shortcut.Arguments = self.run_command()[-36-len(self.filepath):]
        shortcut.WorkingDirectory = wDir
        shortcut.IconLocation = icon
        shortcut.save()
        return path

    def app_startup(self):
        """Add shortcut to startup."""
        file = self.app_desktop()
        folder = r'\Microsoft\Windows\Start Menu\Programs\Startup'
        with open(f'{file[:-12]}file_move.bat', 'w',
                  encoding='1251') as filemove:
            filemove.write(f'''@echo off
chcp 1251>nul
move "{file}" "{os.getenv('appdata')}{folder}\\{file[-12:]}"
del {file[:-12]}file_move.bat 2>nul''')
        os.startfile(f'{file[:-12]}file_move.bat')

    def app_runexit(self):
        """Run and Exit."""
        folder = os.getenv('appdata') + \
            r'\Microsoft\Windows\Start Menu\Programs\Startup'
        for file in os.listdir(folder):
            if len(file) == 12 and file[:-4].isdigit():
                os.startfile(os.path.join(folder, file))
        self.destroy()
        sys.exit()


def main():
    """Basic_function."""
    root = Application()
    root.mainloop()


def test():
    """Test function."""
    print(sys.argv)
    # import locale
    # print(locale.getpreferredencoding())


if __name__ == '__main__':
    if '-run' in sys.argv:
        if getpass('input: ') == get_from_env("PASSW"):
            main()
    elif '-test' in sys.argv:
        test()
    else:
        print('Попытка запуска программы неуспешна')
