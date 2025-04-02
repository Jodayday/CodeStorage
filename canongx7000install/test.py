import json
import tkinter as tk
from tkinter import messagebox, ttk
import subprocess
import os

# 설정 (드라이버 경로, INF 파일 이름)
DRIVER_INF_PATH = r"C:\Users\jojo\Desktop\python_evne\canon\Driver\GX7000P6.inf"
DRIVER_NAME = "Canon GX7000 series"
JSON_PATH = "printers.json"

def install_printer(printer_name, printer_ip):
    try:
        # 드라이버 설치: 이미 설치된 경우를 체크
        try:
            result = subprocess.run(
                ["pnputil", "/add-driver", DRIVER_INF_PATH, "/install"],
                check=True,
                capture_output=True,
                text=True
            )
        except subprocess.CalledProcessError as e:
            error_output = (e.stdout + e.stderr).lower()
            # 에러 메시지에 이미 설치되었다는 내용이 있으면 무시
            if "already installed" in error_output or "이미 설치" in error_output:
                pass
                # print("드라이버가 이미 설치되어 있습니다. 무시하고 진행합니다.")
            else:
                raise  # 다른 오류인 경우 다시 예외 발생

        # TCP/IP 포트 생성 (이미 존재하면 무시)
        try:
            result = subprocess.run([
                "cscript",
                r"C:\Windows\System32\Printing_Admin_Scripts\ko-KR\prnport.vbs",
                "-a", "-r", f"IP_{printer_ip}", "-h", printer_ip, "-o", "raw", "-n", "9100"
            ], capture_output=True, text=True, check=True)
        except subprocess.CalledProcessError as e:
            error_output = (e.stdout or "") + (e.stderr or "")
            # "already exists" 또는 "이미 존재" 메시지가 포함된 경우 예외 무시
            if "already exists" in error_output.lower() or "이미 존재" in error_output.lower():
                pass # print("TCP/IP 포트가 이미 존재합니다. 해당 포트를 사용합니다.")
            else:
                raise

        # 프린터 등록 (이미 포트가 존재해도 상관없음)
        subprocess.run([
            "rundll32", "printui.dll,PrintUIEntry",
            "/if", "/b", printer_name, "/f", DRIVER_INF_PATH,
            "/r", f"IP_{printer_ip}", "/m", DRIVER_NAME
        ], check=True)


        messagebox.showinfo("성공", f"{printer_name} 프린터 설치 완료!")
    except subprocess.CalledProcessError as e:
        messagebox.showerror("오류", f"프린터 설치 중 오류 발생:\n{e}")


# JSON 불러오기
def load_printer_data():
    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        return json.load(f)

# JSON 저장하기
def save_last_school(data, selected_school):
    data["last_selected_school"] = selected_school
    with open(JSON_PATH, 'w', encoding='utf-8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

# 프린터 버튼 로딩
def show_printer_buttons(frame, printer_list):
    for widget in frame.winfo_children():
        widget.destroy()

    for printer in printer_list:
        btn = tk.Button(
            frame,
            text=f"{printer['name']} ({printer['ip']})",
            font=("Arial", 12),
            command=lambda p=printer: install_printer(p['name'], p['ip'])
        )
        btn.pack(pady=5)

# GUI 생성
def create_gui():
    data = load_printer_data()
    schools = list(data["schools"].keys())
    last_school = data.get("last_selected_school", schools[0])

    root = tk.Tk()
    root.title("Canon GX7092 프린터 설치기")
    root.geometry("450x400")

    tk.Label(root, text="학교를 선택하세요:", font=("Arial", 14)).pack(pady=10)

    selected_school = tk.StringVar(value=last_school)
    dropdown = ttk.Combobox(root, textvariable=selected_school, values=schools, state="readonly", font=("Arial", 12))
    dropdown.pack()

    printer_frame = tk.Frame(root)
    printer_frame.pack(pady=20, fill=tk.BOTH, expand=True)

    def on_school_select(event=None):
        school = selected_school.get()
        printer_list = data["schools"][school]
        show_printer_buttons(printer_frame, printer_list)
        save_last_school(data, school)

    dropdown.bind("<<ComboboxSelected>>", on_school_select)
    on_school_select()  # 초기 로딩

    root.mainloop()

# 메인
if __name__ == "__main__":
    if os.path.exists(JSON_PATH):
        create_gui()
    else:
        print("printers.json 파일이 존재하지 않습니다.")
