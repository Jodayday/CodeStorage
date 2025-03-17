import psutil
import subprocess
import ctypes
import sys
import tkinter as tk
from tkinter import messagebox, ttk

# 수동IP 설정하는 프로그램
# 프로그램 순차
# 1.3초마다 연결된 이더넷 인터페이스를 표시함
# 2.ip 입력하면 변경됨
# 3.모든 이더넷 자동변경후 연결된 이더넷에 수동변경함



# 무조건 관리자 권한으로 실행하도록 수정

BASE_IP = "10.131.74."  # 고정 IP의 기본 주소

def get_ethernet_interfaces():
    """이더넷 인터페이스 목록 가져오기"""
    interfaces = []
    net_info = psutil.net_if_addrs()

    for interface in net_info.keys():
        if "Ethernet" in interface or "이더넷" in interface:  # Ethernet 또는 이더넷 포함된 인터페이스만 선택
            interfaces.append(interface)

    return interfaces

def set_dhcp_for_all_ethernet():
    """모든 이더넷 인터페이스를 DHCP로 변경"""
    ethernet_interfaces = get_ethernet_interfaces()

    if not ethernet_interfaces:
        # print("❌ DHCP 변경할 이더넷 인터페이스가 없습니다.")
        return

    # print("🔍 감지된 이더넷 인터페이스:", ethernet_interfaces)

    for interface in ethernet_interfaces:
        try:
            # print(f"🔄 {interface}: DHCP 설정 중...")
            subprocess.run(f'netsh interface ip set address "{interface}" dhcp', shell=True)
            subprocess.run(f'netsh interface ip set dns "{interface}" dhcp', shell=True)
            # print(f"✅ {interface}: DHCP 변경 완료")
        except Exception as e:
            # print(f"❌ 오류 발생 ({interface}): {e}")
            pass

def get_connected_ethernet():
    """현재 연결된 이더넷 인터페이스 반환"""
    stats = psutil.net_if_stats()
    for interface, stat in stats.items():
        if stat.isup and ("Ethernet" in interface or "이더넷" in interface):
            return interface  # 첫 번째 연결된 이더넷 반환
    return None

def update_interface_label():
    """연결된 인터페이스를 라벨에 업데이트"""
    interface = get_connected_ethernet()
    interface_var.set(interface if interface else "연결된 이더넷 없음")
    root.after(3000, update_interface_label)  # 3초마다 갱신

def validate_last_octet(new_value):
    """IP 마지막 옥텟 입력값 검증 (0~255)"""
    if not new_value:
        return True  # 빈 값 허용
    try:
        value = int(new_value)
        return 0 <= value <= 255
    except ValueError:
        return False

def set_static_ip():
    set_dhcp_for_all_ethernet()
    """지정된 이더넷 인터페이스에 고정 IP 설정"""
    interface = get_connected_ethernet()
    if not interface:
        messagebox.showerror("오류", "연결된 이더넷 인터페이스가 없습니다.")
        return

    last_octet = ip_entry.get()
    if not last_octet.isdigit():
        messagebox.showerror("오류", "올바른 마지막 옥텟(0~255)을 입력하세요.")
        return

    ip = BASE_IP + last_octet
    subnet = "255.255.255.0"
    gateway = BASE_IP + "1"
    dns1 = "168.126.63.1"
    dns2 = "168.126.63.2"

    try:
        subprocess.run(f'netsh interface ip set address "{interface}" static {ip} {subnet} {gateway}', shell=True)
        subprocess.run(f'netsh interface ip set dns "{interface}" static {dns1}', shell=True)
        subprocess.run(f'netsh interface ip add dns "{interface}" {dns2} index=2', shell=True)
        messagebox.showinfo("완료", f"{interface}에 {ip} 고정 IP를 설정하였습니다.")
    except Exception as e:
        messagebox.showerror("오류", f"IP 설정 중 오류 발생: {e}")

def is_admin():
    """현재 프로세스가 관리자 권한으로 실행 중인지 확인"""
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    """관리자 권한이 없으면 관리자 모드로 재실행"""
    if not is_admin():
        script = sys.argv[0]
        params = " ".join([f'"{arg}"' for arg in sys.argv[1:]])
        try:
            ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{script}" {params}', None, 1)
            sys.exit(0)  # 현재 프로세스 종료
        except Exception as e:
            messagebox.showerror("오류", f"관리자 권한 실행 실패: {e}")
            sys.exit(1)

run_as_admin()
# =============== GUI 생성 ===============
root = tk.Tk()
root.title("네트워크 설정 도구")
root.geometry("350x250")

# 인터페이스 라벨
ttk.Label(root, text="현재 연결된 이더넷 인터페이스:").pack(pady=5)
interface_var = tk.StringVar()
interface_label = ttk.Label(root, textvariable=interface_var, font=("Arial", 12, "bold"))
interface_label.pack(pady=5)

# 초기 인터페이스 값 설정
update_interface_label()

# 고정 IP 입력 필드 (마지막 옥텟만 입력)
ttk.Label(root, text=f"IP 주소: {BASE_IP}").pack()
validate_cmd = root.register(validate_last_octet)
ip_entry = ttk.Entry(root, width=5, validate="key", validatecommand=(validate_cmd, "%P"))
ip_entry.pack()

# 고정 IP 버튼
static_ip_button = ttk.Button(root, text="고정 IP 설정", command=set_static_ip)
static_ip_button.pack(pady=10)

# 창 실행
root.mainloop()
