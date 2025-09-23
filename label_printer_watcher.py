import os
import sys
import time
import json
import subprocess
import threading
import requests
import zipfile
from datetime import date, datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import tkinter as tk
from tkinter import scrolledtext, messagebox, ttk, filedialog
from queue import Queue
from PIL import Image, ImageDraw, ImageFont, ImageTk
import pystray

# pywin32는 pyinstaller로 빌드 시 자동으로 포함되지 않을 수 있어, 별도 import가 필요할 수 있습니다.
try:
    import win32print
    import win32ui
    import win32con
    import win32gui
    import pywintypes
    from PIL import ImageWin
    import winreg # <--- 추가
except ImportError:
    # 프로그램 실행 중에는 설치할 수 없으므로, 사용자에게 안내 메시지를 표시합니다.
    print("오류: 'pywin32' 모듈을 찾을 수 없습니다. 'pip install pywin32' 명령으로 설치해주세요.")
    win32print = None


# #####################################################################
# 1. 기본 설정 (config.json 파일이 없을 경우 사용)
# #####################################################################
DEFAULT_CONFIG = {
    "REPO_OWNER": "KMTechn",
    "REPO_NAME": "Label_Printer_Watcher",
    "APP_VERSION": "v1.0.1", # 윈도우 시작 시 자동 실행 기능 추가
    "remnant_printer": "",
    "defective_printer": "",
    "remnant_base_folder": "",
    "defective_base_folder": "",
}
CONFIG_FILE = 'config.json'
# #####################################################################

# 설정 관리
def load_config():
    config = DEFAULT_CONFIG.copy()
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            try:
                user_config = json.load(f)
                config.update(user_config)
            except json.JSONDecodeError:
                print(f"오류: '{CONFIG_FILE}' 파일 형식이 잘못되었습니다. 기본 설정으로 복원합니다.")
                save_config(config)
                return config

    config['APP_VERSION'] = DEFAULT_CONFIG['APP_VERSION']
    config['REPO_OWNER'] = DEFAULT_CONFIG['REPO_OWNER']
    config['REPO_NAME'] = DEFAULT_CONFIG['REPO_NAME']
    
    save_config(config)
        
    return config

def save_config(config_data):
    config_to_save = config_data.copy()
    config_to_save.pop('remnant_devmode', None)
    config_to_save.pop('defective_devmode', None)
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(config_to_save, f, indent=4, ensure_ascii=False)

CONFIG = load_config()

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def check_for_updates():
    try:
        owner, repo, version = CONFIG["REPO_OWNER"], CONFIG["REPO_NAME"], CONFIG["APP_VERSION"]
        api_url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
        response = requests.get(api_url, timeout=5)
        response.raise_for_status()
        latest_version = response.json()['tag_name']
        print(f"현재 버전: {version}, 최신 버전: {latest_version}")
        if latest_version.strip().lower() != version.strip().lower():
            for asset in response.json()['assets']:
                if asset['name'].endswith('.zip'):
                    return asset['browser_download_url'], latest_version
        return None, None
    except requests.exceptions.RequestException as e:
        print(f"업데이트 확인 중 오류 발생: {e}")
        return None, None

def download_and_apply_update(url):
    try:
        temp_dir = os.environ.get("TEMP", "C:\\Temp")
        zip_path = os.path.join(temp_dir, "update.zip")
        response = requests.get(url, stream=True, timeout=120)
        response.raise_for_status()
        with open(zip_path, 'wb') as f: f.write(response.content)
        temp_update_folder = os.path.join(temp_dir, "temp_update")
        if os.path.exists(temp_update_folder):
            import shutil
            shutil.rmtree(temp_update_folder)
        with zipfile.ZipFile(zip_path, 'r') as zip_ref:
            zip_ref.extractall(temp_update_folder)
        os.remove(zip_path)
        application_path = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
        updater_script_path = os.path.join(application_path, "updater.bat")
        extracted_content = os.listdir(temp_update_folder)
        new_program_folder_path = os.path.join(temp_update_folder, extracted_content[0]) if len(extracted_content) == 1 and os.path.isdir(os.path.join(temp_update_folder, extracted_content[0])) else temp_update_folder
        with open(updater_script_path, "w", encoding='utf-8') as bat_file:
            bat_file.write(f"""@echo off
chcp 65001 > nul & echo. & echo ========================================================== & echo    프로그램을 업데이트합니다. 이 창을 닫지 마세요. & echo ========================================================== & echo. & echo 잠시 후 프로그램이 자동으로 종료됩니다... & timeout /t 3 /nobreak > nul & taskkill /F /IM "{os.path.basename(sys.executable)}" > nul & echo. & echo 새 파일로 교체합니다... & xcopy "{new_program_folder_path}" "{application_path}" /E /H /C /I /Y > nul & echo. & echo 임시 파일을 삭제합니다... & rmdir /s /q "{temp_update_folder}" & echo. & echo ======================================== & echo    업데이트 완료! & echo ======================================== & echo. & echo 3초 후에 프로그램을 다시 시작합니다. & timeout /t 3 /nobreak > nul & start "" "{os.path.join(application_path, os.path.basename(sys.executable))}" & del "%~f0"
            """)
        subprocess.Popen(updater_script_path, creationflags=subprocess.CREATE_NEW_CONSOLE)
        sys.exit(0)
    except Exception as e:
        messagebox.showerror("업데이트 적용 실패", f"업데이트 적용 중 오류 발생:\n{e}")
        sys.exit(1)

def threaded_update_check():
    print("백그라운드 업데이트 확인 시작...")
    download_url, new_version = check_for_updates()
    if download_url:
        if messagebox.askyesno("업데이트 발견", f"새로운 버전({new_version})이 있습니다.\n지금 업데이트하시겠습니까? (현재 버전: {CONFIG['APP_VERSION']})"):
            download_and_apply_update(download_url)
        else:
            print("사용자가 업데이트를 거부했습니다.")

def print_label(image_path: str, printer_name: str, devmode=None):
    if not win32print:
        print("오류: pywin32 모듈이 없어 인쇄할 수 없습니다.")
        return
    if not os.path.exists(image_path):
        print(f"인쇄 실패: 파일 '{image_path}'를 찾을 수 없습니다.")
        return

    hDC = None
    try:
        print(f"인쇄 시도: '{os.path.basename(image_path)}' -> '{printer_name}'")
        
        if devmode:
            print(" - 시스템 프린터 설정(DEVMODE)을 적용합니다.")
            hDC = win32gui.CreateDC("WINSPOOL", printer_name, devmode)
        else:
            print(" - 기본 프린터 설정을 사용합니다.")
            hDC = win32gui.CreateDC("WINSPOOL", printer_name, None)
        
        hdc = win32ui.CreateDCFromHandle(hDC)

        printable_width = hdc.GetDeviceCaps(win32con.HORZRES)
        printable_height = hdc.GetDeviceCaps(win32con.VERTRES)

        hdc.StartDoc(image_path)
        hdc.StartPage()

        img = Image.open(image_path)
        img_width, img_height = img.size

        img_aspect = img_width / img_height
        printable_aspect = printable_width / printable_height

        if img_aspect > printable_aspect:
            draw_width = printable_width
            draw_height = int(draw_width / img_aspect)
        else:
            draw_height = printable_height
            draw_width = int(draw_height * img_aspect)
        
        dib = ImageWin.Dib(img)
        
        draw_x = (printable_width - draw_width) // 2
        draw_y = (printable_height - draw_height) // 2
        
        dib.draw(hdc.GetHandleOutput(), (draw_x, draw_y, draw_x + draw_width, draw_y + draw_height))

        hdc.EndPage()
        hdc.EndDoc()
        print(f"성공: 인쇄 명령을 전송했습니다.")

    except (win32ui.error, pywintypes.error) as e:
        print(f"오류: 인쇄 중 오류가 발생했습니다. 프린터('{printer_name}') 설정을 확인해주세요.\n{e}")
    except Exception as e:
        print(f"오류: 예기치 않은 인쇄 오류가 발생했습니다.\n{e}")
    finally:
        if hDC:
            win32gui.DeleteDC(hDC)


class LabelPrintHandler(FileSystemEventHandler):
    def __init__(self, printer_name: str, get_devmode_func):
        self.printer_name = printer_name
        self.get_devmode_func = get_devmode_func
        self._last_printed_time = {}

    def on_created(self, event):
        if event.is_directory or not event.src_path.lower().endswith('.png'):
            return
        filepath = event.src_path
        current_time = time.time()
        if self._last_printed_time.get(filepath) and current_time - self._last_printed_time.get(filepath, 0) < 2:
            return
        self._last_printed_time[filepath] = current_time
        
        devmode = self.get_devmode_func()
        threading.Thread(target=print_label, args=(filepath, self.printer_name, devmode), daemon=True).start()

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"라벨 자동 출력기 ({CONFIG['APP_VERSION']})")
        self.geometry("700x700")
        
        # ############### [추가됨] 자동 실행 레지스트리 관련 변수 ###############
        self.app_name = "LabelPrinterWatcher"
        # pyinstaller로 빌드된 .exe 경로 또는 .py 스크립트 경로를 가져옴
        self.app_path = sys.executable if getattr(sys, 'frozen', False) else os.path.abspath(__file__)
        self.reg_key_path = r"Software\Microsoft\Windows\CurrentVersion\Run"

        try:
            logo_path = resource_path('assets/logo.png')
            self.logo_image = ImageTk.PhotoImage(file=logo_path)
            self.iconphoto(True, self.logo_image)
        except Exception as e:
            print(f"윈도우 로고 이미지 로드 실패: {e}")

        self.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.log_queue = Queue()
        self.observer = None
        self.current_watch_date = None
        self.is_running = True
        self.restart_monitoring = threading.Event()
        
        self.remnant_devmode = None
        self.defective_devmode = None

        self.create_widgets()
        self.redirect_stdout()
        self.monitor_thread = threading.Thread(target=self.monitoring_loop, daemon=True)
        self.monitor_thread.start()
        self.after(100, self.process_log_queue)
        self.setup_tray_icon()
        threading.Thread(target=threaded_update_check, daemon=True).start()

    def create_widgets(self):
        self.notebook = ttk.Notebook(self, padding=10)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        log_frame = ttk.Frame(self.notebook)
        self.notebook.add(log_frame, text='실행 로그')
        
        log_label = tk.Label(log_frame, text="실시간 실행 로그", font=("Malgun Gothic", 10, "bold"))
        log_label.pack(anchor="w", pady=(0, 5))
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, font=("Consolas", 10), state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True)

        settings_frame = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(settings_frame, text='인쇄 설정')

        remnant_frame = ttk.LabelFrame(settings_frame, text="잔량 라벨 설정", padding=10)
        remnant_frame.pack(fill=tk.X, pady=5)
        
        tk.Label(remnant_frame, text="프린터 선택:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.remnant_printer_var = tk.StringVar(value=CONFIG.get("remnant_printer", ""))
        tk.Entry(remnant_frame, textvariable=self.remnant_printer_var, width=40).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        tk.Button(remnant_frame, text="찾아보기", command=lambda: self.select_printer(self.remnant_printer_var)).grid(row=0, column=2, padx=5, pady=5)
        
        tk.Button(remnant_frame, text="시스템 프린터 설정", command=lambda: self.open_printer_properties('remnant')).grid(row=0, column=3, padx=5, pady=5)

        tk.Label(remnant_frame, text="감시 폴더:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.remnant_folder_var = tk.StringVar(value=CONFIG.get("remnant_base_folder", ""))
        tk.Entry(remnant_frame, textvariable=self.remnant_folder_var, width=40).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        tk.Button(remnant_frame, text="찾아보기", command=lambda: self.select_folder(self.remnant_folder_var)).grid(row=1, column=2, padx=5, pady=5)
        remnant_frame.columnconfigure(1, weight=1)

        defective_frame = ttk.LabelFrame(settings_frame, text="불량 라벨 설정", padding=10)
        defective_frame.pack(fill=tk.X, pady=5)

        tk.Label(defective_frame, text="프린터 선택:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.defective_printer_var = tk.StringVar(value=CONFIG.get("defective_printer", ""))
        tk.Entry(defective_frame, textvariable=self.defective_printer_var, width=40).grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        tk.Button(defective_frame, text="찾아보기", command=lambda: self.select_printer(self.defective_printer_var)).grid(row=0, column=2, padx=5, pady=5)
        
        tk.Button(defective_frame, text="시스템 프린터 설정", command=lambda: self.open_printer_properties('defective')).grid(row=0, column=3, padx=5, pady=5)

        tk.Label(defective_frame, text="감시 폴더:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.defective_folder_var = tk.StringVar(value=CONFIG.get("defective_base_folder", ""))
        tk.Entry(defective_frame, textvariable=self.defective_folder_var, width=40).grid(row=1, column=1, padx=5, pady=5, sticky="ew")
        tk.Button(defective_frame, text="찾아보기", command=lambda: self.select_folder(self.defective_folder_var)).grid(row=1, column=2, padx=5, pady=5)
        defective_frame.columnconfigure(1, weight=1)
        
        # ############### [추가됨] 자동 실행 설정 프레임 ###############
        startup_frame = ttk.LabelFrame(settings_frame, text="옵션", padding=10)
        startup_frame.pack(fill=tk.X, pady=(15, 5))

        self.startup_var = tk.BooleanVar()
        self.startup_checkbutton = ttk.Checkbutton(startup_frame, text="윈도우 시작 시 자동 실행", variable=self.startup_var, command=self.toggle_startup)
        self.startup_checkbutton.pack(anchor="w")
        self.startup_var.set(self.check_startup_registry())

        button_frame = tk.Frame(settings_frame)
        button_frame.pack(fill=tk.X, pady=20)

        tk.Button(button_frame, text="설정 저장", command=self.save_settings, height=2, bg="#4CAF50", fg="white", font=("Malgun Gothic", 10, "bold")).pack(side=tk.LEFT, expand=True, padx=5)
        tk.Button(button_frame, text="테스트 라벨 생성", command=self.create_test_label, height=2, bg="#2196F3", fg="white", font=("Malgun Gothic", 10, "bold")).pack(side=tk.RIGHT, expand=True, padx=5)

        self.status_var = tk.StringVar(value="초기화 중...")
        status_bar = tk.Label(self, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor='w', padx=5)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    # ############### [추가됨] 자동 실행 레지스트리 관리 함수들 ###############
    def toggle_startup(self):
        if self.startup_var.get():
            self.set_startup_registry()
        else:
            self.remove_startup_registry()

    def set_startup_registry(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.reg_key_path, 0, winreg.KEY_WRITE)
            winreg.SetValueEx(key, self.app_name, 0, winreg.REG_SZ, self.app_path)
            winreg.CloseKey(key)
            print("자동 실행 설정: 레지스트리에 등록되었습니다.")
        except Exception as e:
            messagebox.showerror("오류", f"자동 실행 설정에 실패했습니다:\n{e}")
            print(f"[오류] 자동 실행 레지스트리 등록 실패: {e}")

    def remove_startup_registry(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.reg_key_path, 0, winreg.KEY_WRITE)
            winreg.DeleteValue(key, self.app_name)
            winreg.CloseKey(key)
            print("자동 실행 해제: 레지스트리에서 제거되었습니다.")
        except FileNotFoundError:
            # 이미 없는 경우이므로 정상이므로 무시
            print("자동 실행 해제: 레지스트리에 등록되어 있지 않습니다.")
            pass
        except Exception as e:
            messagebox.showerror("오류", f"자동 실행 해제에 실패했습니다:\n{e}")
            print(f"[오류] 자동 실행 레지스트리 제거 실패: {e}")

    def check_startup_registry(self):
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, self.reg_key_path, 0, winreg.KEY_READ)
            winreg.QueryValueEx(key, self.app_name)
            winreg.CloseKey(key)
            return True
        except FileNotFoundError:
            return False
        except Exception as e:
            print(f"[오류] 자동 실행 상태 확인 실패: {e}")
            return False
            
    def open_printer_properties(self, printer_type):
        if not win32print:
            messagebox.showerror("모듈 오류", "'pywin32'가 설치되지 않았습니다.")
            return

        printer_name = self.remnant_printer_var.get() if printer_type == 'remnant' else self.defective_printer_var.get()

        if not printer_name:
            messagebox.showwarning("프린터 선택 필요", "먼저 프린터를 선택하고 저장해주세요.")
            return

        try:
            PRINTER_DEFAULTS = {"DesiredAccess": win32print.PRINTER_ALL_ACCESS}
            h_printer = win32print.OpenPrinter(printer_name, PRINTER_DEFAULTS)
            
            properties = win32print.GetPrinter(h_printer, 2)
            p_devmode = properties['pDevMode']
            
            result = win32print.DocumentProperties(self.winfo_id(), h_printer, printer_name, p_devmode, p_devmode, win32con.DM_IN_PROMPT | win32con.DM_OUT_BUFFER | win32con.DM_IN_BUFFER)
            
            if result == win32con.IDOK:
                win32print.SetPrinter(h_printer, 2, properties, 0)
                
                if printer_type == 'remnant':
                    self.remnant_devmode = p_devmode
                elif printer_type == 'defective':
                    self.defective_devmode = p_devmode
                
                print(f"프린터({printer_name})의 시스템 기본 설정이 영구적으로 변경되었습니다.")
                messagebox.showinfo("설정 완료", f"'{printer_name}' 프린터의 기본 설정이 **영구적으로** 변경되었습니다.")
            else:
                print("사용자가 프린터 설정 변경을 취소했습니다.")
                messagebox.showinfo("취소", "프린터 설정 변경이 취소되었습니다.")

            win32print.ClosePrinter(h_printer)

        except pywintypes.error as e:
            if e.winerror == 5:
                 messagebox.showerror("권한 오류", f"프린터 설정을 변경할 권한이 없습니다.\n프로그램을 관리자 권한으로 실행해주세요.")
                 print(f"[오류] 프린터 설정 변경 권한 부족: {e}")
            else:
                messagebox.showerror("설정 오류", f"프린터 속성을 여는 중 오류 발생:\n{e}\n프린터 이름이 올바른지 확인해주세요.")
                print(f"[오류] 프린터 속성 열기 실패: {e}")
        except Exception as e:
            messagebox.showerror("알 수 없는 오류", f"예상치 못한 오류 발생:\n{e}")
            print(f"[오류] 프린터 속성 열기 중 예외 발생: {e}")

    def get_current_settings(self):
        return {}

    def select_folder(self, var):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            var.set(folder_selected)

    def get_printers(self):
        if not win32print:
            messagebox.showerror("모듈 오류", "'pywin32' 라이브러리가 설치되지 않았습니다.\n프로그램을 종료하고 'pip install pywin32'를 실행해주세요.")
            return []
        try:
            printers = [printer[2] for printer in win32print.EnumPrinters(2)]
            if not printers:
                print("[정보] 설치된 프린터를 찾을 수 없습니다.")
            return printers
        except Exception as e:
            print(f"[오류] 프린터 목록을 가져오는 데 실패했습니다: {e}")
            messagebox.showerror("프린터 조회 실패", f"프린터 목록을 가져오는 중 오류가 발생했습니다.\n{e}")
            return []

    def select_printer(self, var):
        printers = self.get_printers()
        if not printers:
            messagebox.showinfo("프린터 없음", "시스템에 설치된 프린터가 없습니다.")
            return

        win = tk.Toplevel(self)
        win.title("프린터 선택")
        win.geometry("350x300")
        win.transient(self)
        win.grab_set()

        tk.Label(win, text="설치된 프린터를 선택하세요:", pady=5).pack()

        list_frame = tk.Frame(win)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        scrollbar = tk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(list_frame, yscrollcommand=scrollbar.set)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar.config(command=listbox.yview)

        for printer in printers:
            listbox.insert(tk.END, printer)

        def on_ok():
            selected_indices = listbox.curselection()
            if selected_indices:
                selected_printer = listbox.get(selected_indices[0])
                var.set(selected_printer)
            win.destroy()
        
        def on_double_click(event):
            on_ok()

        listbox.bind("<Double-Button-1>", on_double_click)

        ok_button = tk.Button(win, text="확인", command=on_ok, width=10)
        ok_button.pack(pady=10)

        current_printer = var.get()
        if current_printer in printers:
            idx = printers.index(current_printer)
            listbox.selection_set(idx)
            listbox.see(idx)

        win.wait_window()

    def save_settings(self):
        global CONFIG
        new_config = self.get_current_settings()
        new_config["remnant_printer"] = self.remnant_printer_var.get()
        new_config["remnant_base_folder"] = self.remnant_folder_var.get()
        new_config["defective_printer"] = self.defective_printer_var.get()
        new_config["defective_base_folder"] = self.defective_folder_var.get()
        
        try:
            save_config(new_config)
            CONFIG = load_config()
            messagebox.showinfo("저장 완료", "설정이 성공적으로 저장되었습니다.\n감시자가 재시작됩니다.")
            print("\n[설정 저장] 새로운 설정이 적용되었습니다. 감시자를 재시작합니다.")
            self.restart_monitoring.set()
        except Exception as e:
            messagebox.showerror("저장 실패", f"설정 저장 중 오류가 발생했습니다:\n{e}")
            print(f"[오류] 설정 저장 실패: {e}")

    def create_test_label(self):
        target_folder = self.remnant_folder_var.get()
        if not target_folder or not os.path.isdir(target_folder):
            messagebox.showwarning("폴더 오류", "잔량 라벨 감시 폴더가 올바르게 설정되지 않았습니다.\n설정 탭에서 폴더를 지정해주세요.")
            return

        try:
            today_str = date.today().strftime('%Y-%m-%d')
            today_folder = os.path.join(target_folder, today_str)
            os.makedirs(today_folder, exist_ok=True)

            width, height = 400, 200
            img = Image.new('RGB', (width, height), color = 'white')
            d = ImageDraw.Draw(img)

            try:
                font = ImageFont.truetype("malgun.ttf", 20)
            except IOError:
                font = ImageFont.load_default()

            text = "테스트 라벨입니다."
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            d.text((20,20), "--- 샘플 인쇄 ---", fill=(0,0,0), font=font)
            d.text((20,60), text, fill=(0,0,0), font=font)
            d.text((20,100), f"생성 시간: {timestamp}", fill=(0,0,0), font=font)
            d.rectangle([(5,5), (width-5, height-5)], outline ="black", width=2)

            filename = f"test_label_{int(time.time())}.png"
            filepath = os.path.join(today_folder, filename)
            img.save(filepath)

            messagebox.showinfo("생성 완료", f"테스트 라벨이 생성되었습니다.\n경로: {filepath}\n\n잠시 후 설정된 프린터로 자동 인쇄됩니다.")
            print(f"[테스트] 샘플 라벨 생성 완료: {filepath}")

        except Exception as e:
            messagebox.showerror("생성 실패", f"테스트 라벨 생성 중 오류가 발생했습니다:\n{e}")
            print(f"[오류] 테스트 라벨 생성 실패: {e}")


    def add_log(self, message):
        self.log_text.config(state='normal')
        self.log_text.insert(tk.END, message)
        self.log_text.see(tk.END)
        self.log_text.config(state='disabled')
    
    def process_log_queue(self):
        while not self.log_queue.empty():
            message = self.log_queue.get_nowait()
            self.add_log(message)
        self.after(100, self.process_log_queue)

    def redirect_stdout(self):
        class StdoutRedirector:
            def __init__(self, queue):
                self.queue = queue
            def write(self, str):
                self.queue.put(str)
            def flush(self):
                pass
        sys.stdout = StdoutRedirector(self.log_queue)
        sys.stderr = StdoutRedirector(self.log_queue)

    def monitoring_loop(self):
        print("--- 자동 라벨 출력 프로그램 시작 ---\n")
        while self.is_running:
            today = date.today()
            if today != self.current_watch_date or self.restart_monitoring.is_set():
                if self.observer and self.observer.is_alive():
                    print(f"\n감시자를 재시작합니다. (사유: {'설정변경' if self.restart_monitoring.is_set() else '날짜변경'})")
                    self.observer.stop()
                    self.observer.join()
                
                self.restart_monitoring.clear()
                self.observer = Observer()
                self.current_watch_date = today
                today_str = today.strftime('%Y-%m-%d')
                
                print(f"\n[{today_str}] 날짜의 폴더 감시를 설정합니다.")
                remnant_base, remnant_printer = CONFIG.get("remnant_base_folder"), CONFIG.get("remnant_printer")
                if remnant_base and remnant_printer and os.path.isdir(remnant_base):
                    remnant_path_today = os.path.join(remnant_base, today_str)
                    os.makedirs(remnant_path_today, exist_ok=True)
                    print(f" - 잔량 폴더 감시 중: {remnant_path_today} -> [{remnant_printer}]")
                    self.observer.schedule(LabelPrintHandler(remnant_printer, lambda: self.remnant_devmode), remnant_path_today, recursive=False)
                else:
                    print(f" - 잔량 폴더 설정이 올바르지 않아 감시를 시작할 수 없습니다.")

                defective_base, defective_printer = CONFIG.get("defective_base_folder"), CONFIG.get("defective_printer")
                if defective_base and defective_printer and os.path.isdir(defective_base):
                    defective_path_today = os.path.join(defective_base, today_str)
                    os.makedirs(defective_path_today, exist_ok=True)
                    print(f" - 불량 폴더 감시 중: {defective_path_today} -> [{defective_printer}]")
                    self.observer.schedule(LabelPrintHandler(defective_printer, lambda: self.defective_devmode), defective_path_today, recursive=False)
                else:
                    print(f" - 불량 폴더 설정이 올바르지 않아 감시를 시작할 수 없습니다.")

                if self.observer.emitters:
                    self.observer.start()
                    self.status_var.set(f"모니터링 중... (감시 날짜: {today_str})")
                    print("모니터링 시작...\n")
                else:
                    self.status_var.set("오류: 감시할 폴더가 설정되지 않았습니다.")
                    print("\n감시할 폴더가 설정되지 않았습니다. '인쇄 설정' 탭에서 설정을 확인해주세요.")
            
            for _ in range(60):
                if not self.is_running or self.restart_monitoring.is_set():
                    break
                time.sleep(1)

    def setup_tray_icon(self):
        try:
            logo_path = resource_path('assets/logo.png')
            image = Image.open(logo_path)
        except Exception as e:
            print(f"트레이 아이콘 로고 로드 실패: {e}")
            image = Image.new('RGB', (64, 64), 'blue')
        menu = (pystray.MenuItem('열기', self.show_window, default=True),
                pystray.MenuItem('종료', self.on_closing))
        self.tray_icon = pystray.Icon("name", image, "라벨 자동 출력기", menu)
        threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def show_window(self):
        self.deiconify()
        self.lift()
        self.focus_force()

    def on_closing(self):
        if messagebox.askyesno("종료 확인", "프로그램을 종료하시겠습니까?"):
            self.quit_app()

    def quit_app(self):
        print("프로그램 종료 중...")
        self.is_running = False
        if self.observer and self.observer.is_alive():
            self.observer.stop()
            self.observer.join()
        self.tray_icon.stop()
        self.destroy()

if __name__ == "__main__":
    app = App()
    app.mainloop()