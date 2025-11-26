import sys
import os
import time
import clr
import math
import random
from threading import Thread

# ==========================================
#  虛擬鍵碼對照表 (Virtual Key Codes) - 保持不變
# ==========================================
VK = {
    'BACKSPACE': 0x08, 'TAB': 0x09, 'ENTER': 0x0D, 'SHIFT': 0x10,
    'CTRL': 0x11, 'ALT': 0x12, 'ESC': 0x1B, 'SPACE': 0x20,
    'LEFT': 0x25, 'UP': 0x26, 'RIGHT': 0x27, 'DOWN': 0x28,
    'DELETE': 0x2E,
    '0': 0x30, '1': 0x31, '2': 0x32, '3': 0x33, '4': 0x34,
    '5': 0x35, '6': 0x36, '7': 0x37, '8': 0x38, '9': 0x39,
    'A': 0x41, 'B': 0x42, 'C': 0x43, 'D': 0x44, 'E': 0x45,
    'F': 0x46, 'G': 0x47, 'H': 0x48, 'I': 0x49, 'J': 0x4A,
    'K': 0x4B, 'L': 0x4C, 'M': 0x4D, 'N': 0x4E, 'O': 0x4F,
    'P': 0x50, 'Q': 0x51, 'R': 0x52, 'S': 0x53, 'T': 0x54,
    'U': 0x55, 'V': 0x56, 'W': 0x57, 'X': 0x58, 'Y': 0x59, 'Z': 0x5A,
    'LWIN': 0x5B, 'RWIN': 0x5C, 
    'F1': 0x70, 'F2': 0x71, 'F3': 0x72, 'F4': 0x73, 'F5': 0x74,
    'F6': 0x75, 'F7': 0x76, 'F8': 0x77, 'F9': 0x78, 'F10': 0x79,
}

class RdpController:
    def __init__(self, dll_path=None):
        """ 初始化 RDP 控制器 """
        if dll_path is None:
            # 調整為你的 DLL 路徑
            base_dir = os.path.dirname(os.path.abspath(__file__))
            dll_path = os.path.join(base_dir, 'RdpClientBridge', 'bin', 'Debug', 'RdpClientBridge.dll')
            # 如果上面路徑找不到，嘗試 Release
            if not os.path.exists(dll_path):
                 dll_path = os.path.join(base_dir, 'RdpClientBridge', 'bin', 'Release', 'RdpClientBridge.dll')
        
        if not os.path.exists(dll_path):
            # 最後嘗試直接在腳本目錄找
             dll_path = os.path.join(base_dir, 'RdpClientBridge.dll')

        if not os.path.exists(dll_path):
             raise FileNotFoundError(f"DLL 遺失: {dll_path}")

        self.dll_dir = os.path.dirname(dll_path)
        self.dll_path = dll_path
        self.rdp_instance = None
        self.is_running = False
        self.thread = None
        self._load_references()

    def _load_references(self):
        """載入必要的 .NET DLL"""
        clr.AddReference("System")
        clr.AddReference("System.Windows.Forms")
        try:
            clr.AddReference(os.path.join(self.dll_dir, "Interop.MSTSCLib.dll"))
            clr.AddReference(os.path.join(self.dll_dir, "AxInterop.MSTSCLib.dll"))
        except:
            pass
        clr.AddReference(self.dll_path)
        
        import System
        from System.Windows.Forms import Application
        self._System = System
        self._Application = Application
        self.assembly = System.Reflection.Assembly.Load("RdpClientBridge")
        self.RDPConnectionType = self.assembly.GetType("RdpClientBridge.RDPConnection")

    def _sta_thread_func(self, server, username, password, port, width, height):
        try:
            self.rdp_instance = self._System.Activator.CreateInstance(self.RDPConnectionType)
            self.rdp_instance.Server = server
            self.rdp_instance.Username = username
            self.rdp_instance.Password = password
            self.rdp_instance.Port = port
            self.rdp_instance.Width = width
            self.rdp_instance.Height = height
            self.rdp_instance.Connect()
            self._Application.Run(self.rdp_instance)
        except Exception as e:
            print(f"[RdpController Error] {e}")
        finally:
            self.is_running = False
            self.rdp_instance = None

    def connect(self, server, username, password, port=3389, width=1024, height=768, background=False):
        if self.is_running: return
        self.is_running = True
        self.thread = Thread(target=self._sta_thread_func, args=(server, username, password, port, width, height))
        self.thread.setDaemon(True)
        self.thread.start()

        # 等待初始化
        timeout = 15
        start_time = time.time()
        while self.rdp_instance is None:
            if time.time() - start_time > timeout:
                self.is_running = False
                raise TimeoutError("RDP 實例初始化超時。")
            time.sleep(0.1)
        
        if background:
            time.sleep(1) # 等視窗出來再隱藏
            self.hide_window()

    def close(self):
        if self.rdp_instance:
            try:
                # 必須在 UI Thread 關閉，這裡用 Invoke 或簡單的 Application.Exit
                # 這裡簡單標記，實際關閉依賴 FormClosing
                pass 
            except: pass
        self.is_running = False

    # --- 操作 API ---
    def click(self, x, y):
        if self.rdp_instance: self.rdp_instance.SendMouseClickBackground(int(x), int(y))
    
    def right_click(self, x, y):
        if self.rdp_instance: self.rdp_instance.SendMouseRightClickBackground(int(x), int(y))

    def press_key(self, key_code_or_char):
        if not self.rdp_instance: return
        vk_code = 0
        if isinstance(key_code_or_char, int):
            vk_code = key_code_or_char
        elif isinstance(key_code_or_char, str):
            key = key_code_or_char.upper()
            if key in VK: vk_code = VK[key]
            elif key.isdigit(): vk_code = int(key)
            elif len(key) == 1: vk_code = ord(key)
        if vk_code > 0: self.rdp_instance.SendKeyBackground(vk_code)

    def type_text(self, text, interval=0.05):
        for char in text:
            if char == ' ': self.press_key('SPACE')
            else: self.press_key(char)
            time.sleep(interval)

    def hide_window(self):
        if self.rdp_instance: self.rdp_instance.MoveToBackground()

    def show_window(self):
        if self.rdp_instance: self.rdp_instance.RestoreWindow()

    # --- 原子操作 ---
    def mouse_down(self, x, y):
        if self.rdp_instance: self.rdp_instance.SendMouseDownBackground(int(x), int(y))
    def mouse_up(self, x, y):
        if self.rdp_instance: self.rdp_instance.SendMouseUpBackground(int(x), int(y))
    def mouse_move(self, x, y, is_drag=False):
        if self.rdp_instance: self.rdp_instance.SendMouseMoveBackground(int(x), int(y), is_drag)

    # --- 拖曳算法 (簡化版) ---
    def drag_drop(self, x1, y1, x2, y2):
        self.mouse_move(x1, y1)
        time.sleep(0.1)
        self.mouse_down(x1, y1)
        time.sleep(0.1)
        # 簡單插值移動
        steps = 20
        dx = (x2 - x1) / steps
        dy = (y2 - y1) / steps
        for i in range(steps):
            self.mouse_move(int(x1 + dx * i), int(y1 + dy * i), True)
            time.sleep(0.01)
        self.mouse_move(x2, y2, True)
        time.sleep(0.1)
        self.mouse_up(x2, y2)

# ==========================================
#  多連線管理器 (Multi-Session Manager)
# ==========================================
class MultiRdpManager:
    def __init__(self):
        self.sessions = {} # { 'id': RdpController }
        self.current_id = None

    def add_session(self, session_id, ip, user, pwd, port=3389, hide=False):
        if session_id in self.sessions:
            print(f"Error: ID '{session_id}' already exists.")
            return False
        
        try:
            print(f"[{session_id}] Connecting to {ip}...")
            bot = RdpController()
            bot.connect(ip, user, pwd, port, background=hide)
            self.sessions[session_id] = bot
            
            # 如果是第一個連線，自動設為當前
            if self.current_id is None:
                self.current_id = session_id
            
            print(f"[{session_id}] Connected.")
            return True
        except Exception as e:
            print(f"[{session_id}] Connection Failed: {e}")
            return False

    def get_current(self):
        if not self.current_id: return None
        return self.sessions.get(self.current_id)

    def switch_session(self, session_id):
        if session_id in self.sessions:
            self.current_id = session_id
            print(f"Switched to session: {session_id}")
        else:
            print(f"Session '{session_id}' not found.")

    def close_all(self):
        for sid, bot in self.sessions.items():
            print(f"Closing {sid}...")
            bot.close()
        self.sessions.clear()

# ==========================================
#  多工互動模式 (Interactive Mode)
# ==========================================
def interactive_mode():
    print("=== Multi-RDP Manager CLI ===")
    manager = MultiRdpManager()
    
    # 預設參數 (方便測試)
    def_ip = "127.0.0.1"
    def_user = "Administrator"
    def_pwd = "password"

    print("指令說明:")
    print("  new <id> [ip] [user] [pwd]  - 建立新連線")
    print("  use <id>                    - 切換控制目標")
    print("  list                        - 列出所有連線")
    print("  hide / show                 - 隱藏/顯示 當前視窗")
    print("  click <x> <y>               - 點擊 (當前視窗)")
    print("  type <text>                 - 輸入文字 (當前視窗)")
    print("  drag <x1> <y1> <x2> <y2>    - 拖曳")
    print("  exit                        - 結束所有連線並退出")
    print("-----------------------------")

    while True:
        try:
            # 提示當前控制的 Session ID
            prompt_id = manager.current_id if manager.current_id else "NoSession"
            cmd_line = input(f"[{prompt_id}]> ").strip()
            if not cmd_line: continue

            parts = cmd_line.split()
            cmd = parts[0].lower()

            # --- 全域管理指令 ---
            if cmd == 'exit' or cmd == 'q':
                manager.close_all()
                break
            
            elif cmd == 'list':
                print("Active Sessions:")
                for sid, bot in manager.sessions.items():
                    active = "Running" if bot.is_running else "Closed"
                    marker = "*" if sid == manager.current_id else " "
                    print(f" {marker} [{sid}] {active}")

            elif cmd == 'use':
                if len(parts) < 2: print("Usage: use <id>")
                else: manager.switch_session(parts[1])

            elif cmd == 'new':
                # new session1 192.168.1.10 admin 1234
                if len(parts) < 2:
                    print("Usage: new <id> [ip] [user] [pwd]")
                    continue
                
                sid = parts[1]
                ip = parts[2] if len(parts) > 2 else def_ip
                user = parts[3] if len(parts) > 3 else def_user
                pwd = parts[4] if len(parts) > 4 else def_pwd
                
                manager.add_session(sid, ip, user, pwd)

            # --- 針對當前 Session 的指令 ---
            else:
                bot = manager.get_current()
                if not bot or not bot.is_running:
                    print("Error: No active session selected.")
                    continue

                if cmd == 'hide':
                    bot.hide_window()
                elif cmd == 'show':
                    bot.show_window()
                elif cmd == 'click':
                    if len(parts) == 3: bot.click(parts[1], parts[2])
                elif cmd == 'rclick':
                    if len(parts) == 3: bot.right_click(parts[1], parts[2])
                elif cmd == 'type':
                    if len(parts) > 1: bot.type_text(" ".join(parts[1:]))
                elif cmd == 'key':
                    if len(parts) == 2: bot.press_key(parts[1])
                elif cmd == 'drag':
                    if len(parts) == 5: bot.drag_drop(int(parts[1]), int(parts[2]), int(parts[3]), int(parts[4]))
                elif cmd == 'wait':
                    if len(parts) == 2: time.sleep(float(parts[1]))
                else:
                    print("Unknown command.")

        except Exception as e:
            print(f"Error: {e}")

    os._exit(0)

if __name__ == "__main__":
    interactive_mode()