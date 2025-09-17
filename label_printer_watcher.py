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

# #####################################################################
# 자동 업데이트 설정 (Auto-Updater Configuration)
# #####################################################################
REPO_OWNER = "KMTechn"
REPO_NAME = "Label_Printer_Watcher"
APP_VERSION = "v1.1.0"

# (자동 업데이트 함수들은 이전과 동일하므로 생략... 만약 전체 코드가 필요하시면 다시 요청해주세요)
def check_for_updates():
    try:
        api_url = f"https://api.github.com/repos/{REPO_OWNER}/{REPO_NAME}/releases/latest"
        response = requests.get(api_url, timeout=5)
        response.raise_for_status()
        latest_version = response.json()['tag_name']
        print(f"현재 버전: {APP_VERSION}, 최신 버전: {latest_version}")
        if latest_version.strip().lower() != APP_VERSION.strip().lower():
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
        with open(zip_path, 'wb') as f:
            for chunk in response.iter_content(chunk_size=8192):
                f.write(chunk)
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
chcp 65001 > nul
echo.
echo ==========================================================
echo    프로그램을 업데이트합니다. 이 창을 닫지 마세요.
echo ==========================================================
echo.
echo 잠시 후 프로그램이 자동으로 종료됩니다...
timeout /t 3 /nobreak > nul
taskkill /F /IM "{os.path.basename(sys.executable)}" > nul
echo.
echo 새 파일로 교체합니다...
xcopy "{new_program_folder_path}" "{application_path}" /E /H /C /I /Y > nul
echo.
echo 임시 파일을 삭제합니다...
rmdir /s /q "{temp_update_folder}"
echo.
echo ========================================
echo    업데이트 완료!
echo ========================================
echo.
echo 3초 후에 프로그램을 다시 시작합니다.
timeout /t 3 /nobreak > nul
start "" "{os.path.join(application_path, os.path.basename(sys.executable))}"
del "%~f0"
            """)
        subprocess.Popen(updater_script_path, creationflags=subprocess.CREATE_NEW_CONSOLE)
        sys.exit(0)
    except Exception as e:
        print(f"업데이트 적용 실패: {e}")
        time.sleep(5)
        sys.exit(1)

def threaded_update_check():
    import tkinter as tk
    from tkinter import messagebox
    print("백그라운드 업데이트 확인 시작...")
    download_url, new_version = check_for_updates()
    if download_url:
        root = tk.Tk()
        root.withdraw()
        if messagebox.askyesno("업데이트 발견", f"새로운 버전({new_version})이 있습니다.\n지금 업데이트하시겠습니까? (현재 버전: {APP_VERSION})"):
            root.destroy()
            download_and_apply_update(download_url)
        else:
            print("사용자가 업데이트를 거부했습니다.")
            root.destroy()

# #####################################################################
# 설정 및 인쇄 로직 (Configuration and Printing Logic)
# #####################################################################

def get_config_path():
    """실행 파일(.exe) 또는 스크립트(.py)가 있는 폴더를 기준으로 설정 파일의 전체 경로를 반환합니다."""
    # PyInstaller로 빌드된 .exe 파일에서 실행될 때의 경로
    if getattr(sys, 'frozen', False):
        base_path = os.path.dirname(sys.executable)
    # .py 스크립트로 직접 실행될 때의 경로
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, 'printer_config.json')

def load_or_create_config():
    """설정 파일을 로드하거나, 없으면 스크립트와 같은 위치에 기본값으로 생성합니다."""
    config_path = get_config_path() # [수정] 설정 파일의 정확한 경로를 가져옴
    
    if not os.path.exists(config_path):
        print(f"'{config_path}' 파일이 없어 기본값으로 생성합니다.")
        default_config = {
            "remnant_printer": "잔량 라벨 프린터 이름을 여기에 입력",
            "remnant_base_folder": "C:\\Sync\\labels\\remnant_labels",
            "defective_printer": "불량 라벨 프린터 이름을 여기에 입력",
            "defective_base_folder": "C:\\Sync\\labels\\defective_labels"
        }
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, indent=4, ensure_ascii=False)
        print(f"-> 중요: 생성된 '{os.path.basename(config_path)}' 파일을 열어 프린터 이름을 정확하게 수정해주세요.")
        return default_config
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"설정 파일 로드 중 오류 발생: {e}")
        return None

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
        if self._last_printed_time.get(filepath) and current_time - self._last_printed_time[filepath] < 2:
            return
        self._last_printed_time[filepath] = current_time
        threading.Thread(target=print_label, args=(filepath, self.printer_name), daemon=True).start()

# #####################################################################
# 메인 실행 로직 (Main Execution Logic)
# #####################################################################
if __name__ == "__main__":
    threading.Thread(target=threaded_update_check, daemon=True).start()

    print("--- 자동 라벨 출력 프로그램 시작 ---")
    
    config = load_or_create_config() # [수정] 수정된 함수 호출
    if not config:
        input("오류로 인해 프로그램을 종료합니다. Enter 키를 눌러주세요...")
        sys.exit(1)

    observer = None
    current_watch_date = None

    try:
        while True:
            today = date.today()
            if today != current_watch_date:
                if observer and observer.is_alive():
                    print(f"\n날짜가 변경되었습니다. ({current_watch_date} -> {today}) 감시자를 재시작합니다.")
                    observer.stop()
                    observer.join()

                observer = Observer()
                current_watch_date = today
                today_str = today.strftime('%Y-%m-%d')
                print(f"\n[{today_str}] 날짜의 폴더 감시를 설정합니다.")

                # 잔량 폴더 감시 설정
                remnant_base = config.get("remnant_base_folder")
                remnant_printer = config.get("remnant_printer")
                if remnant_base and remnant_printer and "여기에 입력" not in remnant_printer:
                    remnant_path_today = os.path.join(remnant_base, today_str)
                    os.makedirs(remnant_path_today, exist_ok=True)
                    print(f" - 잔량 폴더 감시 중: {remnant_path_today}")
                    observer.schedule(LabelPrintHandler(remnant_printer), remnant_path_today, recursive=False)

                # 불량 폴더 감시 설정
                defective_base = config.get("defective_base_folder")
                defective_printer = config.get("defective_printer")
                if defective_base and defective_printer and "여기에 입력" not in defective_printer:
                    defective_path_today = os.path.join(defective_base, today_str)
                    os.makedirs(defective_path_today, exist_ok=True)
                    print(f" - 불량 폴더 감시 중: {defective_path_today}")
                    observer.schedule(LabelPrintHandler(defective_printer), defective_path_today, recursive=False)
                
                if observer.emitters:
                    observer.start()
                    print("모니터링 시작...")
                else:
                    config_path = get_config_path() # [수정] 정확한 설정 파일 경로 안내
                    print(f"\n감시할 폴더가 설정되지 않았습니다. '{os.path.basename(config_path)}' 파일의 프린터 이름이 올바른지 확인해주세요.")

            time.sleep(60)

    except KeyboardInterrupt:
        print("\n프로그램 종료 신호를 받았습니다. 감시를 중지합니다.")
    finally:
        if observer and observer.is_alive():
            observer.stop()
            observer.join()
        print("프로그램이 안전하게 종료되었습니다.")