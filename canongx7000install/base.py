import json
import tkinter as tk
from tkinter import messagebox, ttk
import subprocess
import os
import threading

# 설정 (드라이버 경로, INF 파일 이름)
DRIVER_INF_PATH = r"C:\Users\admin\Desktop\canon_install\Driver\GX7000P6.inf"
DRIVER_NAME = "Canon GX7000 series"
JSON_PATH = "printers.json"

def is_printer_installed(printer_name):
    """
    wmic 명령을 이용하여 현재 설치된 프린터 목록에서
    printer_name 이 이미 존재하는지 확인합니다.
    """
    try:
        result = subprocess.run(
            ["wmic", "printer", "get", "name"],
            capture_output=True, text=True, check=True
        )
        installed_printers = result.stdout
        # 설치된 프린터 목록에 대상 이름이 포함되어 있는지 확인
        if printer_name in installed_printers:
            return True
    except subprocess.CalledProcessError:
        pass
    return False

def install_printer(printer_name, printer_ip):
    """
    프린터 설치 작업을 진행합니다.
    이미 설치된 경우 관련 예외를 발생시키지 않고 진행합니다.
    """
    # 드라이버 설치: 이미 설치된 경우를 체크
    try:
        subprocess.run(
            ["pnputil", "/add-driver", DRIVER_INF_PATH, "/install"],
            check=True,
            capture_output=True,
            text=True
        )
    except subprocess.CalledProcessError as e:
        error_output = (e.stdout or "") + (e.stderr or "")
        if "already installed" in error_output.lower() or "이미 설치" in error_output.lower():
            print("드라이버가 이미 설치되어 있습니다. 무시하고 진행합니다.")
        else:
            raise

    # TCP/IP 포트 생성 (이미 존재하면 무시)
    try:
        subprocess.run([
            "cscript",
            r"C:\Windows\System32\Printing_Admin_Scripts\ko-KR\prnport.vbs",
            "-a", "-r", f"IP_{printer_ip}", "-h", printer_ip, "-o", "raw", "-n", "9100"
        ], capture_output=True, text=True, check=True)
    except subprocess.CalledProcessError as e:
        error_output = (e.stdout or "") + (e.stderr or "")
        if "already exists" in error_output.lower() or "이미 존재" in error_output.lower():
            print("TCP/IP 포트가 이미 존재합니다. 해당 포트를 사용합니다.")
        else:
            raise

    # 프린터 등록
    subprocess.run([
        "rundll32", "printui.dll,PrintUIEntry",
        "/if", "/b", printer_name, "/f", DRIVER_INF_PATH,
        "/r", f"IP_{printer_ip}", "/m", DRIVER_NAME
    ], check=True)

def threaded_install(printer_name, printer_ip, status_label, root):
    """
    별도 스레드에서 프린터 설치를 진행합니다.
    설치 시작 전에 이미 설치된 경우 이를 체크하여
    진행하지 않고 '이미 설치' 메시지를 표시합니다.
    """
    # 설치 전에 해당 프린터가 이미 설치되어 있는지 확인
    if is_printer_installed(printer_name):
        root.after(0, lambda: status_label.config(text=f"{printer_name}: 이미 설치되어 있습니다."))
        root.after(0, lambda: messagebox.showinfo("정보", f"{printer_name} 는 이미 설치되어 있습니다."))
        return

    try:
        root.after(0, lambda: status_label.config(text=f"{printer_name}: 설치 중..."))
        install_printer(printer_name, printer_ip)
    except Exception as e:
        root.after(0, lambda: status_label.config(text=f"{printer_name}: 설치 실패"))
        root.after(0, lambda: messagebox.showerror("오류", f"{printer_name} 프린터 설치 중 오류 발생:\n{e}"))
    else:
        root.after(0, lambda: status_label.config(text=f"{printer_name}: 설치 완료"))
        root.after(0, lambda: messagebox.showinfo("성공", f"{printer_name} 프린터 설치 완료!"))

# JSON 불러오기
def load_printer_data():
    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

# JSON 저장하기 (마지막 선택한 학교 업데이트)
def save_last_school(data, selected_school):
    data["last_selected_school"] = selected_school
    with open(JSON_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

# 프린터 버튼 로딩 (버튼 클릭 시 별도 스레드에서 설치)
def show_printer_buttons(frame, printer_list, status_label, root):
    for widget in frame.winfo_children():
        widget.destroy()

    for printer in printer_list:
        btn = tk.Button(
            frame,
            text=f"{printer['name']} ({printer['ip']})",
            font=("Arial", 12),
            command=lambda p=printer: threading.Thread(
                target=threaded_install, 
                args=(p['name'], p['ip'], status_label, root)
            ).start()
        )
        btn.pack(pady=5)



# GUI 생성
def create_gui():
    data = load_printer_data()
    schools = list(data["schools"].keys())
    last_school = data.get("last_selected_school", schools[0])

    root = tk.Tk()
    root.title("Canon GX7092 프린터 설치기")
    root.geometry("450x450")

    tk.Label(root, text="학교를 선택하세요:", font=("Arial", 14)).pack(pady=10)

    selected_school = tk.StringVar(value=last_school)
    dropdown = ttk.Combobox(root, textvariable=selected_school, values=schools, state="readonly", font=("Arial", 12))
    dropdown.pack()

    printer_frame = tk.Frame(root)
    printer_frame.pack(pady=20, fill=tk.BOTH, expand=True)

    # 설치 진행 상태를 보여줄 레이블
    status_label = tk.Label(root, text="", font=("Arial", 12))
    status_label.pack(pady=10)

    def on_school_select(event=None):
        school = selected_school.get()
        printer_list = data["schools"][school]
        show_printer_buttons(printer_frame, printer_list, status_label, root)
        save_last_school(data, school)

    dropdown.bind("<<ComboboxSelected>>", on_school_select)
    on_school_select()  # 초기 로딩

    root.mainloop()

# 메인 실행
if __name__ == "__main__":
    if os.path.exists(JSON_PATH):
        create_gui()
    else:
        print("printers.json 파일이 존재하지 않습니다.")
