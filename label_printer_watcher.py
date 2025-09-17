import os
import sys
import time
import json
import subprocess
import threading
import requests
import zipfile
from datetime import date
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import tkinter as tk
from tkinter import scrolledtext, messagebox
from queue import Queue
from PIL import Image, ImageTk
import pystray

# #####################################################################
# 1. 사용자 설정 (이 부분을 직접 수정하세요)
# #####################################################################
CONFIG = {
    "REPO_OWNER": "KMTechn",
    "REPO_NAME": "Label_Printer_Watcher",
    "APP_VERSION": "v1.0.0", # 경로 수정 버전
    "remnant_printer": "Beeprt BY-482BT_무선",
    "defective_printer": "Beeprt BY-482BT_무선",
    "remnant_base_folder": "C:\\Sync\\labels\\remnant_labels",
    "defective_base_folder": "C:\\Sync\\labels\\defective_labels"
}
# #####################################################################

# [수정됨] .py 또는 .exe 파일의 실제 위치를 기준으로 경로를 찾는 함수
def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller는 임시 폴더를 만들고 _MEIPASS에 경로를 저장합니다.
        base_path = sys._MEIPASS
    except Exception:
        # 일반 .py 스크립트로 실행될 때는 파일 자신의 위치를 기준으로 합니다.
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# (자동 업데이트 및 인쇄 관련 함수들은 이전과 동일)
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

def print_label(image_path: str, printer_name: str):
    if not os.path.exists(image_path):
        print(f"인쇄 실패: 파일 '{image_path}'를 찾을 수 없습니다.")
        return
    try:
        print(f"인쇄 시도: '{os.path.basename(image_path)}' -> '{printer_name}'")
        subprocess.run(["mspaint", "/p", image_path, printer_name], check=True, shell=True)
        print(f"성공: 인쇄 명령을 전송했습니다.")
    except Exception as e:
        print(f"오류: 인쇄 중 오류가 발생했습니다. 프린터('{printer_name}')가 연결되어 있는지 확인해주세요.\n{e}")

class LabelPrintHandler(FileSystemEventHandler):
    def __init__(self, printer_name: str):
        self.printer_name = printer_name
        self._last_printed_time = {}
    def on_created(self, event):
        if event.is_directory or not event.src_path.lower().endswith('.png'):
            return
        filepath = event.src.path
        current_time = time.time()
        if self._last_printed_time.get(filepath) and current_time - self._last_printed_time.get(filepath, 0) < 2:
            return
        self._last_printed_time[filepath] = current_time
        threading.Thread(target=print_label, args=(filepath, self.printer_name), daemon=True).start()

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"라벨 자동 출력기 ({CONFIG['APP_VERSION']})")
        self.geometry("700x400")
        
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

        self.create_widgets()
        self.redirect_stdout()
        self.monitor_thread = threading.Thread(target=self.monitoring_loop, daemon=True)
        self.monitor_thread.start()
        self.after(100, self.process_log_queue)
        self.setup_tray_icon()
        threading.Thread(target=threaded_update_check, daemon=True).start()

    def create_widgets(self):
        main_frame = tk.Frame(self, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        log_label = tk.Label(main_frame, text="실행 로그", font=("Malgun Gothic", 10, "bold"))
        log_label.pack(anchor="w")
        self.log_text = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, font=("Consolas", 10), state='disabled')
        self.log_text.pack(fill=tk.BOTH, expand=True, pady=(5, 10))
        self.status_var = tk.StringVar(value="초기화 중...")
        status_bar = tk.Label(self, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor='w', padx=5)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

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
            if today != self.current_watch_date:
                if self.observer and self.observer.is_alive():
                    print(f"\n날짜가 변경되었습니다. ({self.current_watch_date} -> {today}) 감시자를 재시작합니다.")
                    self.observer.stop()
                    self.observer.join()
                self.observer = Observer()
                self.current_watch_date = today
                today_str = today.strftime('%Y-%m-%d')
                print(f"\n[{today_str}] 날짜의 폴더 감시를 설정합니다.")
                remnant_base, remnant_printer = CONFIG.get("remnant_base_folder"), CONFIG.get("remnant_printer")
                if remnant_base and remnant_printer:
                    remnant_path_today = os.path.join(remnant_base, today_str)
                    os.makedirs(remnant_path_today, exist_ok=True)
                    print(f" - 잔량 폴더 감시 중: {remnant_path_today} -> [{remnant_printer}]")
                    self.observer.schedule(LabelPrintHandler(remnant_printer), remnant_path_today, recursive=False)
                defective_base, defective_printer = CONFIG.get("defective_base_folder"), CONFIG.get("defective_printer")
                if defective_base and defective_printer:
                    defective_path_today = os.path.join(defective_base, today_str)
                    os.makedirs(defective_path_today, exist_ok=True)
                    print(f" - 불량 폴더 감시 중: {defective_path_today} -> [{defective_printer}]")
                    self.observer.schedule(LabelPrintHandler(defective_printer), defective_path_today, recursive=False)
                if self.observer.emitters:
                    self.observer.start()
                    self.status_var.set(f"모니터링 중... (감시 날짜: {today_str})")
                    print("모니터링 시작...\n")
                else:
                    self.status_var.set("오류: 감시할 폴더가 설정되지 않았습니다.")
                    print("\n감시할 폴더가 설정되지 않았습니다. 코드 상단의 CONFIG 설정을 확인해주세요.")
            time.sleep(60)

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