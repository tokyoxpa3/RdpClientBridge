import sys
import os
import clr

dll_path = os.path.join(os.path.dirname(__file__), 'RdpClientBridge.dll')

if not os.path.exists(dll_path):
    print(f"!! 錯誤：RDP 客戶端組件 (DLL) 遺失。\n請確認路徑是否正確，或 C# 專案是否已成功編譯: {dll_path}")
    sys.exit(1)

# 取得 DLL 所在的目錄
dll_dir = os.path.dirname(dll_path)

# 載入 .NET 系統組件，用於執行緒和 UI
clr.AddReference("System")
clr.AddReference("System.Threading")
clr.AddReference("System.Windows.Forms")

# 【重要修正】明確地預先載入依賴項
# 這有助於 .NET 執行環境在載入主 DLL 之前解析其所需的元件
try:
    clr.AddReference(os.path.join(dll_dir, "Interop.MSTSCLib.dll"))
    clr.AddReference(os.path.join(dll_dir, "AxInterop.MSTSCLib.dll"))
except Exception as e:
    print(f"!! 警告：載入依賴項 DLL 時發生問題。如果後續出錯，這可能是原因所在。")
    print(f"   詳細資訊: {e}")

# 載入主 DLL
clr.AddReference(dll_path)

# --- 診斷步驟：使用 .NET 反射來檢查 Assembly ---
try:
    import System
    # `clr.AddReference` 已經將 Assembly 載入到 AppDomain
    # 現在我們使用完整的名稱來載入它以便進行檢查
    # 注意：'RdpClientBridge' 是 AssemblyName，不是 DLL 的檔名
    assembly = System.Reflection.Assembly.Load("RdpClientBridge")
    
    print("\n--- 偵錯診斷：開始檢查 RdpClientBridge.dll ---")
    print(f"成功載入 Assembly: {assembly.FullName}")
    print("\n尋找其中所有公開的類別 (Exported Types):")
    
    found_types = False
    # 使用 try-except 包裹 GetExportedTypes()，因為如果依賴項有問題，它會拋出例外
    try:
        exported_types = assembly.GetExportedTypes()
    except System.Reflection.ReflectionTypeLoadException as ex:
        print("  - (!!) 錯誤：GetExportedTypes() 失敗。這強烈表示有依賴項載入失敗。")
        print("    LoaderExceptions:")
        for loader_ex in ex.LoaderExceptions:
            print(f"      - {loader_ex.Message}")
        exported_types = ex.Types # 顯示成功載入的部分

    for t in exported_types:
        if t is not None:
            print(f"  - 找到: {t.FullName}")
            found_types = True
        
    if not found_types:
        print("  - (!!) 未找到任何成功載入的公開類別。")
    
    print("\n--- 診斷結束 ---\n")

    # 如果診斷程式碼本身就出錯，下面的 import 也會失敗
    print("診斷後，再次嘗試導入 RDPConnection...")
    # The most robust way to get the type is directly from the assembly
    # object we loaded during the diagnostic phase. This bypasses Python's
    # sometimes unreliable import hooks for .NET.
    RDPConnection = assembly.GetType("RdpClientBridge.RDPConnection")
    print(">>> 成功！已成功獲取 RDPConnection 類別的引用。")

    from System.Threading import Thread, ThreadStart, ApartmentState
    from System.Windows.Forms import MessageBox, MessageBoxButtons, MessageBoxIcon, Application

except Exception as e:
    print(f"\n!! 在診斷或導入過程中發生錯誤。\n   詳細錯誤: {e}")
    # 印出堆疊追蹤以便更深入分析
    import traceback
    traceback.print_exc()
    sys.exit(1)


# --- 2. STA 執行緒中的 RDP 連線邏輯 ---

def rdp_sta_thread_action(server, username, password, port, width, height):
    """
    此函數將在一個獨立的 STA 執行緒中執行，專門用於處理 RDP 連線。
    """
    try:
        # a. 實例化 C# 的 RDPConnection 類別 (現在它是一個 Form)
        print("在 STA 執行緒中：正在建立 RDPConnection 物件...")
        # 【修正】使用 System.Activator.CreateInstance 來實例化從 reflection 獲取的 .NET 型別
        rdp_client = System.Activator.CreateInstance(RDPConnection)

        # b. 設定連線屬性
        rdp_client.Server = server
        rdp_client.Username = username
        rdp_client.Password = password
        rdp_client.Port = port
        rdp_client.Width = width
        rdp_client.Height = height
        rdp_client.Fullscreen = False # 可在此處修改

        # c. 呼叫 Connect 方法來應用屬性
        print("在 STA 執行緒中：正在呼叫 Connect() 來應用設定...")
        rdp_client.Connect()
        
        # d. 運行視窗的訊息循環
        print("在 STA 執行緒中：正在呼叫 Application.Run() 來顯示視窗...")
        Application.Run(rdp_client)
        print("在 STA 執行緒中：Application.Run() 已結束 (RDP 視窗已關閉)。")

    except Exception as e:
        # 在執行緒內處理任何 .NET 拋出的例外
        error_message = f"RDP 執行緒內部發生錯誤: {e}"
        print(f"!! {error_message}")
        # 我們也可以在這裡用 .NET 的 MessageBox 來顯示錯誤，因為我們在 UI 執行緒上
        MessageBox.Show(error_message, "Python RDP Connector Error", MessageBoxButtons.OK, MessageBoxIcon.Error)


# --- 3. 主程式邏輯 ---

def main():
    print("RDP 連線器 - Python 與 .NET 整合工具")
    print("=" * 50)

    response = input("\n使用預設資料連線？(Y/n): ").strip().lower()
    if response in ['n', 'no', '否', '不']:
        print("請輸入 RDP 連線資訊：")
        default_server = "127.0.0.2"
        default_port = 3389
        server = input(f"伺服器地址 (預設: {default_server}): ").strip() or default_server
        port_input = input(f"連接埠 (預設: {default_port}): ").strip()
    else:
        server = "127.0.0.2"
        port_input = "3389"
        username = ""
        password = ""
    
    try:
        port = int(port_input) if port_input else default_port
    except ValueError:
        port = default_port
        print(f"無效的連接埠，使用預設值: {port}")
    
    print("\n" + "=" * 50)
    print(f"準備連接到: {server}:{port}")
    print(f"使用者: {username}")
    print("\n注意：")
    print("- 將會開啟一個新的 RDP 連線視窗。")
    print("- 關閉 RDP 視窗即可中斷連線並結束程式。")
    print("\n正在啟動 RDP 連線...")
    try:
        # 建立一個 .NET 執行緒，並將其 Action 設定為我們的 RDP 函數
        # 使用 lambda 來傳遞參數
        sta_thread = Thread(ThreadStart(lambda: rdp_sta_thread_action(server, username, password, port, 800, 600)))
        
        # 這是最關鍵的一步：將執行緒設定為單一執行緒單元 (STA)
        sta_thread.SetApartmentState(ApartmentState.STA)
        
        # 啟動執行緒
        sta_thread.Start()

        # 等待 RDP 執行緒結束 (即 RDP 視窗被關閉)
        sta_thread.Join()
        
        print("\nRDP 連線程序已結束。\n")

    except Exception as e:
        print(f"\n--- 啟動 RDP 連線時發生未預期的主錯誤 ---")
        print(f"錯誤訊息: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
