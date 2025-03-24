# -*- coding: utf-8 -*-
import json
import os
import ctypes
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from psutil import net_if_stats
from subprocess import run as subprocess_run  # `run` 대신 `subprocess_run` 사용

# JSON 파일 경로
CONFIG_FILE = "config.json"

# 학교별 IP 데이터를 저장할 변수
school_data = {}
current_school = ""
dns1 = ""
dns2 = ""

# JSON 파일 경로 설정 함수
def get_config_path():
    """실행 환경에 따라 JSON 파일 경로를 반환"""
    if getattr(sys, 'frozen', False):  # PyInstaller로 패키징된 실행 파일인지 확인
        base_path = sys._MEIPASS  # 실행 파일이 위치한 임시 경로
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))  # 일반적인 Python 실행 경로

    return os.path.join(base_path, "config.json")

def run_command(command,shell=False):
    """subprocess.run()을 최적화하여 실행 속도 개선"""
    try:
        subprocess_run(command, check=True)
    except Exception as e:
        pass
        # messagebox.showerror("오류", f"명령 실행 중 오류 발생: {e}")

def load_config():
    """JSON 파일에서 학교별 IP, DNS, 마지막 선택한 학교 불러오기"""
    global school_data, current_school, dns1, dns2
    config_path = get_config_path()  # JSON 파일 경로 가져오기

    if os.path.exists(config_path):
        with open(config_path, "r", encoding="utf-8") as file:
            config = json.load(file)
            school_data = config.get("schools", {})
            current_school = config.get("last_school", "")
            dns1 = config.get("dns1", "8.8.8.8")
            dns2 = config.get("dns2", "8.8.4.4")
    else:
        messagebox.showerror("오류", f"{config_path} 파일을 찾을 수 없습니다.")

def save_config():
    """JSON 파일을 비동기적으로 저장 (경로 문제 해결)"""
    def save():
        config_path = get_config_path()  # 올바른 JSON 경로 가져오기
        config = {
            "last_school": current_school,
            "dns1": dns1,
            "dns2": dns2,
            "schools": school_data
        }
        try:
            with open(config_path, "w", encoding="utf-8") as file:
                json.dump(config, file, indent=4, ensure_ascii=False)
        except Exception as e:
            messagebox.showerror("오류", f"설정 저장 중 오류 발생: {e}")

    threading.Thread(target=save, daemon=True).start()

def get_connected_ethernet():
    """현재 연결된 이더넷 인터페이스 반환"""
    return next((iface for iface, stat in net_if_stats().items() if stat.isup and ("Ethernet" in iface or "이더넷" in iface)), None)

def set_dhcp_for_all_ethernet():
    """'이더넷'이 포함된 모든 인터페이스를 DHCP로 변경"""
    # net_if_stats()로 반환된 인터페이스 중 이름에 '이더넷'이 포함된 것만 처리
    for interface in net_if_stats():
        if "이더넷" not in interface:
            continue
        try:
            run_command(["netsh", "interface", "ip", "set", "address", interface, "dhcp"], shell=False)
            run_command(["netsh", "interface", "ip", "set", "dns", interface, "dhcp"], shell=False)
        except Exception:
            print(f"DHCP 오류 발생: {interface}")
            # 오류 발생 시에도 다음 인터페이스로 진행

def async_set_dhcp_for_all_ethernet():
    """set_dhcp_for_all_ethernet 함수를 백그라운드 스레드에서 실행"""
    threading.Thread(target=set_dhcp_for_all_ethernet, daemon=True).start()
def set_static_ip(event=None):
    """모든 이더넷이 DHCP로 설정된 후 지정된 이더넷 인터페이스에 고정 IP 설정 (비동기 실행)"""

    def configure_static_ip():
        # 1. 모든 이더넷 인터페이스를 DHCP로 설정
        set_dhcp_for_all_ethernet()
        
        # 2. DHCP 설정이 완료된 후, 연결된 이더넷 인터페이스 가져오기
        interface = get_connected_ethernet()
        if not interface:
            messagebox.showerror("오류", "연결된 이더넷 인터페이스가 없습니다.")
            return

        # 3. 학교 선택 및 입력 값 검증
        if not current_school or current_school not in school_data:
            messagebox.showerror("오류", "학교를 먼저 선택하세요.")
            return

        last_octet = ip_entry.get()
        if not last_octet.isdigit() or not (0 <= int(last_octet) <= 255):
            messagebox.showerror("오류", "올바른 마지막 옥텟(0~255)을 입력하세요.")
            return

        base_ip = school_data[current_school]
        ip = base_ip + last_octet
        subnet = "255.255.255.0"
        gateway = base_ip + "1"

        # 4. 정적 IP 및 DNS 설정 적용
        try:
            run_command(["netsh", "interface", "ip", "set", "address", interface, "static", ip, subnet, gateway],shell=False)
            run_command(["netsh", "interface", "ip", "set", "dns", interface, "static", dns1],shell=False)
            run_command(["netsh", "interface", "ip", "add", "dns", interface, dns2, "index=2"],shell=False)
            messagebox.showinfo("완료", f"{interface}에 {ip} 고정 IP를 설정하였습니다.")
        except Exception as e:
            print(f"고정 오류 발생: {interface}")
            
            # messagebox.showerror("오류", f"IP 설정 중 오류 발생: {e}")

    # configure_static_ip() 함수 전체를 백그라운드 스레드에서 실행하여 GUI가 멈추지 않도록 함
    threading.Thread(target=configure_static_ip, daemon=True).start()

def change_school():
    """학교 변경 시 비밀번호 입력 후 변경"""
    password_window = tk.Toplevel(root)
    password_window.title("암호 입력")
    password_window.geometry("250x120")
    password_window.resizable(False, False)  # 창 크기 고정

    # 프로그램 창(root) 위에 표시되도록 위치 조정
    root_x = root.winfo_x()
    root_y = root.winfo_y()
    password_window.geometry(f"+{root_x + 50}+{root_y + 50}")  # root 창 기준으로 위치 설정

    ttk.Label(password_window, text="암호를 입력하세요:").pack(pady=5)

    password_var = tk.StringVar()
    password_entry = ttk.Entry(password_window, textvariable=password_var, show="*")
    password_entry.pack(pady=5)
    password_entry.focus_set()  # 창이 뜨면 자동으로 입력창에 포커스

    def check_password(event=None):
        """입력된 암호 확인"""
        if password_var.get() == "3967":
            password_window.destroy()
            school_selection_window()  # 암호가 맞으면 학교 변경 창 열기
        else:
            messagebox.showerror("오류", "잘못된 암호입니다.")
            password_entry.delete(0, tk.END)  # 입력 필드 초기화

    password_entry.bind("<Return>", check_password)  # 엔터 키 입력 처리
    ttk.Button(password_window, text="확인", command=check_password).pack(pady=5)


def school_selection_window():
    """학교 선택을 위한 팝업 윈도우"""
    selection_window = tk.Toplevel(root)
    selection_window.title("학교 선택")
    selection_window.geometry("250x150")
    selection_window.resizable(False, False)  # 창 크기 고정
    # 프로그램 창(root) 위에 표시되도록 위치 조정
    root_x = root.winfo_x()
    root_y = root.winfo_y()
    selection_window.geometry(f"+{root_x + 50}+{root_y + 50}")  # root 창 기준으로 위치 설정

    ttk.Label(selection_window, text="학교 선택:").pack(pady=5)
    
    school_var = tk.StringVar()
    school_dropdown = ttk.Combobox(selection_window, values=list(school_data.keys()), state="readonly", textvariable=school_var)
    school_dropdown.pack(pady=5)
    school_dropdown.focus_set()

    def update_selected_school(event=None):
        global current_school
        selected_school = school_var.get()
        if selected_school in school_data:
            current_school = selected_school
            save_config()
            update_school_info()
            # messagebox.showinfo("학교 변경 완료", f"{selected_school}로 변경되었습니다.")
        selection_window.destroy()

    school_dropdown.bind("<Return>", update_selected_school)  # 엔터키로 변경
    ttk.Button(selection_window, text="변경", command=update_selected_school).pack(pady=5)

def update_school_info():
    """학교 정보 업데이트"""
    if current_school in school_data:
        school_name_var.set(current_school)  # 학교명 업데이트
        base_ip_var.set(school_data[current_school])  # 기본 IP 업데이트
    else:
        school_name_var.set("학교 선택 안됨")
        base_ip_var.set("")

def validate_last_octet(new_value):
    """IP 마지막 옥텟 입력값 검증 (0~255)"""
    return new_value.isdigit() and 0 <= int(new_value) <= 255 if new_value else True


def update_interface_label():
    """연결된 인터페이스를 백그라운드 스레드에서 업데이트"""
    def update():
        interface = get_connected_ethernet()
        interface_var.set(interface if interface else "연결된 이더넷 없음")
        root.after(2500, lambda: threading.Thread(target=update).start())  # 비동기 실행

    threading.Thread(target=update, daemon=True).start()


def change_pc_name(new_name):
    """PowerShell을 사용하여 컴퓨터 이름 변경"""
    try:
        print(f"컴퓨터 이름이 '{new_name}'(으)로 변경되었습니다.")
    except Exception as e:
        print(f"오류 발생: {e}")


def is_admin():
    """현재 프로세스가 관리자 권한으로 실행 중인지 확인"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """관리자 권한이 없으면 관리자 모드로 재실행"""
    if is_admin():
        return
    script = sys.argv[0]
    params = " ".join([f'"{arg}"' for arg in sys.argv[1:]])
    try:
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script}" {params}', None, 1)
        sys.exit(0)
    except Exception as e:
        messagebox.showerror("오류", f"관리자 권한 실행 실패: {e}")
        sys.exit(1)

run_as_admin()

# =============== GUI 생성 ===============
root = tk.Tk()
root.title("네트워크 설정 도구")
root.geometry("300x230")
root.resizable(False, False)  # 창 크기 변경 불가능 (가로, 세로 고정)

# 메인 프레임 생성 (모든 위젯을 포함할 컨테이너)
main_frame = ttk.Frame(root)
main_frame.pack(fill="both", expand=True, padx=10, pady=10)

# 학교별 IP 불러오기
load_config()

# 메뉴바 생성
menubar = tk.Menu(root)
school_menu = tk.Menu(menubar, tearoff=0)
school_menu.add_command(label="학교 변경", command=change_school)
menubar.add_cascade(label="설정", menu=school_menu)
root.config(menu=menubar)

# 인터페이스 라벨
ttk.Label(main_frame, text="현재 연결된 이더넷 인터페이스:").pack(pady=5)
interface_var = tk.StringVar()
interface_label = ttk.Label(main_frame, textvariable=interface_var, font=("Arial", 12, "bold"))
interface_label.pack(pady=5)

update_interface_label()

# 선택된 학교명 및 기본 IP 표시용 프레임 생성
school_info_frame = ttk.Frame(main_frame)
school_info_frame.pack(pady=5, fill="x")

# 선택된 학교명 표시
ttk.Label(school_info_frame, text="선택된 학교:").pack(side="left", padx=5)
school_name_var = tk.StringVar()
school_name_label = ttk.Label(school_info_frame, textvariable=school_name_var, font=("Arial", 10, "bold"))
school_name_label.pack(side="left", padx=5)

# 기본 IP 주소 표시용 프레임 생성
ip_info_frame = ttk.Frame(main_frame)
ip_info_frame.pack(pady=5, fill="x")

# 기본 IP 주소 라벨
ttk.Label(ip_info_frame, text="기본 IP 주소:").pack(side="left", padx=5)
base_ip_var = tk.StringVar()  # 기본 IP 저장 변수
ttk.Label(ip_info_frame, textvariable=base_ip_var, font=("Arial", 10)).pack(side="left", padx=5)

# IP 입력 필드
validate_cmd = root.register(validate_last_octet)
ip_entry = ttk.Entry(main_frame, width=5, validate="key", validatecommand=(validate_cmd, "%P"))
ip_entry.pack()
ip_entry.bind("<Return>", set_static_ip)

# 버튼들을 담을 프레임 생성 (메인 프레임 내)
button_frame = ttk.Frame(main_frame)
button_frame.pack(pady=10)

# 고정 IP 설정 버튼 (왼쪽)
static_ip_button = ttk.Button(button_frame, text="고정 IP 설정", command=set_static_ip)
static_ip_button.pack(side="left", padx=5)

# 자동 IP 설정 버튼 (오른쪽)
auto_ip_button = ttk.Button(button_frame, text="자동IP설정", command=set_dhcp_for_all_ethernet)
auto_ip_button.pack(side="left", padx=5)

# 마지막 선택한 학교 정보 반영
update_school_info()

root.mainloop()
