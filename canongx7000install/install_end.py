import json
import tkinter as tk
from tkinter import messagebox, ttk
import subprocess
import os,sys
import threading

# 설정 (드라이버 경로, INF 파일 이름)
DRIVER_INF_PATH = r"C:\Users\admin\Desktop\canon_install\Driver\GX7000P6.inf"
DRIVER_NAME = "Canon GX7000 series"
JSON_PATH = "printers.json"
current_school = ""

#inno 로 설치시 파일 위치..
DRIVER_INF_PATH = r"C:\Program Files (x86)\Canon7000Installer\Driver\GX7000P6.inf"
JSON_PATH = r"C:\Program Files (x86)\Canon7000Installer\printers.json"

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

def load_printer_data():
    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

def save_last_school(data, selected_school):
    data["last_selected_school"] = selected_school
    with open(JSON_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

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

def school_selection_window():
    """학교 선택을 위한 팝업 윈도우"""
    selection_window = tk.Toplevel(root)
    selection_window.title("학교 선택")
    selection_window.geometry("250x150")
    selection_window.resizable(False, False)
    root_x = root.winfo_x()
    root_y = root.winfo_y()
    selection_window.geometry(f"+{root_x + 50}+{root_y + 50}")

    ttk.Label(selection_window, text="학교 선택:").pack(pady=5)
    
    school_var = tk.StringVar()
    school_dropdown = ttk.Combobox(selection_window, values=schools, state="readonly", textvariable=school_var)
    school_dropdown.pack(pady=5)
    school_dropdown.focus_set()

    def update_selected_school(event=None):
        global current_school, data
        selected = school_var.get()
        if selected in schools:
            current_school = selected
            data["last_selected_school"] = selected
            save_last_school(data, selected)
            selected_school.set(selected)  # 메인 UI에 반영
            on_school_select()  # 프린터 목록 갱신
        selection_window.destroy()

    school_dropdown.bind("<Return>", update_selected_school)
    ttk.Button(selection_window, text="변경", command=update_selected_school).pack(pady=5)

def change_school():
    """학교 변경 시 비밀번호 입력 후 변경"""
    password_window = tk.Toplevel(root)
    password_window.title("암호 입력")
    password_window.geometry("250x120")
    password_window.resizable(False, False)
    root_x = root.winfo_x()
    root_y = root.winfo_y()
    password_window.geometry(f"+{root_x + 50}+{root_y + 50}")

    ttk.Label(password_window, text="암호를 입력하세요:").pack(pady=5)

    password_var = tk.StringVar()
    password_entry = ttk.Entry(password_window, textvariable=password_var, show="*")
    password_entry.pack(pady=5)
    password_entry.focus_set()

    def check_password(event=None):
        if password_var.get() == "3967":
            password_window.destroy()
            school_selection_window()  # 암호가 맞으면 학교 변경 창 열기
        else:
            messagebox.showerror("오류", "잘못된 암호입니다.")
            password_entry.delete(0, tk.END)

    password_entry.bind("<Return>", check_password)
    ttk.Button(password_window, text="확인", command=check_password).pack(pady=5)

# 메인 GUI 생성
data = load_printer_data()
schools = list(data["schools"].keys())
last_school = data.get("last_selected_school", schools[0])

root = tk.Tk()
root.title("Canon GX7092 프린터 설치기")
root.geometry("450x450")

# 메인 UI의 학교명은 StringVar로 관리하여 변경시 자동 갱신
selected_school = tk.StringVar(value=last_school)
tk.Label(root, textvariable=selected_school, font=("Arial", 14)).pack(pady=10)

# 메뉴바 생성
menubar = tk.Menu(root)
school_menu = tk.Menu(menubar, tearoff=0)
school_menu.add_command(label="학교 변경", command=change_school)
menubar.add_cascade(label="설정", menu=school_menu)
root.config(menu=menubar)

printer_frame = tk.Frame(root)
printer_frame.pack(pady=20, fill=tk.BOTH, expand=True)

status_label = tk.Label(root, text="", font=("Arial", 12))
status_label.pack(pady=10)

def on_school_select(event=None):
    school = selected_school.get()
    printer_list = data["schools"][school]
    show_printer_buttons(printer_frame, printer_list, status_label, root)
    save_last_school(data, school)

on_school_select()  # 초기 로딩

root.mainloop()
