import os
import time
import clr
from threading import Thread

# ==========================================
#  虛擬鍵碼對照表 (Virtual Key Codes)
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
    'LWIN': 0x5B, 'RWIN': 0x5C, # Windows 鍵
    'F1': 0x70, 'F2': 0x71, 'F3': 0x72, 'F4': 0x73, 'F5': 0x74,
    'F6': 0x75, 'F7': 0x76, 'F8': 0x77, 'F9': 0x78, 'F10': 0x79,
}

class RdpController:
    def __init__(self, dll_path=None):
        """
        初始化 RDP 控制器
        """
        if dll_path is None:
            # 預設路徑：相對於當前腳本的位置
            base_dir = os.path.dirname(os.path.abspath(__file__))
            dll_path = os.path.join(base_dir ,'RdpClientBridge.dll')
        
        if not os.path.exists(dll_path):
            raise FileNotFoundError(f"DLL 遺失: {dll_path}")

        self.dll_dir = os.path.dirname(dll_path)
        self.dll_path = dll_path
        
        # 內部狀態
        self.rdp_instance = None
        self.is_running = False
        self.thread = None
        
        # 載入元件
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
        
        # 反射載入類別
        self.assembly = System.Reflection.Assembly.Load("RdpClientBridge")
        self.RDPConnectionType = self.assembly.GetType("RdpClientBridge.RDPConnection")

    def _sta_thread_func(self, server, username, password, port, width, height):
        """STA 執行緒：負責 UI 訊息循環"""
        try:
            self.rdp_instance = self._System.Activator.CreateInstance(self.RDPConnectionType)
            
            self.rdp_instance.Server = server
            self.rdp_instance.Username = username
            self.rdp_instance.Password = password
            self.rdp_instance.Port = port
            self.rdp_instance.Width = width
            self.rdp_instance.Height = height
            
            self.rdp_instance.Connect()
            
            # 啟動 Windows 訊息循環 (這會阻塞直到視窗關閉)
            self._Application.Run(self.rdp_instance)
            
        except Exception as e:
            print(f"[RdpController Error] {e}")
        finally:
            self.is_running = False
            self.rdp_instance = None

    def connect(self, server, username, password, port=3389, width=1024, height=768, background=False):
        """啟動 RDP 連線"""
        if self.is_running:
            print("RDP 已經在執行中。")
            return

        self.is_running = True
        self.thread = Thread(target=self._sta_thread_func, args=(server, username, password, port, width, height))
        self.thread.setDaemon(True)
        self.thread.start()

        # 等待初始化
        print("等待 RDP 連線初始化...", end="", flush=True)
        timeout = 15
        start_time = time.time()
        while self.rdp_instance is None:
            if time.time() - start_time > timeout:
                self.is_running = False
                raise TimeoutError("\nRDP 實例初始化超時。")
            time.sleep(0.5)
            print(".", end="", flush=True)
        
        print(" 完成。")
        time.sleep(1) # 緩衝
        
        if background:
            self.hide_window()

    def close(self):
        """標記結束"""
        self.is_running = False

    # ==========================
    #      自動化 API
    # ==========================

    def click(self, x, y):
        """發送滑鼠點擊 (x, y)"""
        if self.rdp_instance:
            self.rdp_instance.SendMouseClickBackground(int(x), int(y))

    def press_key(self, key_code_or_char):
        """
        發送按鍵
        參數: 'ENTER', 'A', 'LWIN' 或整數鍵碼 65
        """
        if not self.rdp_instance: return

        vk_code = 0
        if isinstance(key_code_or_char, int):
            vk_code = key_code_or_char
        elif isinstance(key_code_or_char, str):
            key = key_code_or_char.upper()
            if key in VK:
                vk_code = VK[key]
            elif key.isdigit():
                 vk_code = int(key)
            elif len(key) == 1:
                vk_code = ord(key)
            else:
                print(f"[Warn] 未知按鍵: {key}")
                return
        
        if vk_code > 0:
            self.rdp_instance.SendKeyBackground(vk_code)

    def type_text(self, text, interval=0.05):
        """輸入字串"""
        for char in text:
            if char == ' ':
                self.press_key('SPACE')
            else:
                self.press_key(char)
            time.sleep(interval)

    def hide_window(self):
        """隱藏視窗 (背景模式)"""
        if self.rdp_instance:
            self.rdp_instance.MoveToBackground()

    def show_window(self):
        """顯示視窗"""
        if self.rdp_instance:
            self.rdp_instance.RestoreWindow()

    def sleep(self, seconds):
        """等待秒數"""
        time.sleep(seconds)


# ==========================================
#  互動式調試模式 (Interactive Mode)
# ==========================================
def interactive_mode():
    print("=== RDP 自動化調試模式 (Interactive Mode) ===")
    
    try:
        bot = RdpController()
    except Exception as e:
        print(f"初始化錯誤: {e}")
        return

    # 1. 取得連線資訊
    def_server = "127.0.0.2"
    def_user = "Administrator"
    def_pwd = "password"

    srv = input(f"Server IP [{def_server}]: ").strip() or def_server
    usr = input(f"Username  [{def_user}]: ").strip() or def_user
    pwd = input(f"Password  [******]: ").strip() or def_pwd

    print("\n正在連線...")
    try:
        bot.connect(srv, usr, pwd)
        print("連線成功！視窗已啟動。")
    except Exception as e:
        print(f"連線失敗: {e}")
        return

    # 2. 顯示指令說明
    print("\n=== 指令列表 ===")
    print("  click:<x>,<y>   - 滑鼠點擊 (例: click:100,200)")
    print("  key:<name>      - 按鍵 (例: key:ENTER 或 key:A 或 key:LWIN)")
    print("  type:<text>     - 輸入文字 (例: type:notepad)")
    print("  hide            - 隱藏視窗 (背景模式)")
    print("  show            - 顯示視窗")
    print("  wait:<sec>      - 等待秒數")
    print("  q               - 退出程式")
    print("================\n")

    # 3. 指令迴圈
    while bot.is_running:
        try:
            user_input = input("RDP-Debug> ").strip()
            if not user_input: continue

            # 解析指令
            parts = user_input.split(':', 1)
            cmd = parts[0].lower()
            val = parts[1].strip() if len(parts) > 1 else ""

            if cmd == 'q' or cmd == 'exit':
                break
            
            elif cmd == 'hide':
                bot.hide_window()
                print("-> Window Hidden")
            
            elif cmd == 'show':
                bot.show_window()
                print("-> Window Restored")

            elif cmd == 'click':
                # 格式 click:100,200
                coords = val.split(',')
                if len(coords) == 2:
                    bot.click(coords[0], coords[1])
                    print(f"-> Clicked at ({coords[0]}, {coords[1]})")
                else:
                    print("! 格式錯誤: click:x,y")

            elif cmd == 'key':
                # 格式 key:ENTER 或 key:65
                if val.isdigit():
                    bot.press_key(int(val))
                else:
                    bot.press_key(val)
                print(f"-> Key Sent: {val}")

            elif cmd == 'type':
                # 格式 type:hello world
                bot.type_text(val)
                print(f"-> Typed: {val}")

            elif cmd == 'wait':
                sec = float(val)
                print(f"Waiting {sec}s...")
                time.sleep(sec)

            else:
                print("! 未知指令")

        except KeyboardInterrupt:
            print("\n! 中斷 (再次 Ctrl+C 退出)")
        except Exception as e:
            print(f"! 執行錯誤: {e}")

    bot.close()
    print("程式結束。")
    os._exit(0)

if __name__ == "__main__":
    interactive_mode()