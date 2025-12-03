# RDP Client Bridge

一個用於橋接 Python 和 C# RDP (Remote Desktop Protocol) 客戶端功能的專案，提供簡便的 RDP 連線管理。

## 概述

RDP Client Bridge 是一個 .NET C# 程式庫，它封裝了 Microsoft RDP 客戶端控制項，並提供 COM 介面和 .NET 組件，讓 Python 應用程式可以輕鬆建立和管理 RDP 連線。

## 功能特色

- 透過 C# Windows Forms 應用程式建立 RDP 連線
- 支援多種連線參數配置（伺服器、使用者名稱、密碼、連接埠、解析度等）
- 提供 Python 集成支援（透過 pythonnet 或 COM 介面）
- 支援全螢幕模式和自訂解析度
- 內建進階 RDP 設定（快取持久性、加速器傳遞等）
- 支援後台鍵鼠操作（無需視窗焦點即可發送鍵盤和滑鼠輸入）
- 支援滑鼠右鍵點擊功能
- 支援滑鼠拖曳操作（滑鼠按下、滑鼠放開、滑鼠移動）

## 系統需求

### Windows 系統需求
- Windows 7/8/10/11 或 Windows Server 2012 及更新版本
- .NET Framework 4.8 或更高版本
- Microsoft Remote Desktop Client (內建於 Windows)

### Python 環境需求（如需 Python 整合）
- Python 3.7 或更高版本（64 位元版本）
- pythonnet 套件（`pip install pythonnet`）
- 或 comtypes 套件（`pip install comtypes`）

## 安裝步驟

### 方法一：直接使用 C# 組件

1. 編譯專案以生成 `RdpClientBridge.dll`
2. 在需要使用 RDP 功能的應用程式中引用此 DLL

### 方法二：Python 整合

#### 使用 pythonnet

1. 安裝 pythonnet：
   ```bash
   pip install pythonnet
   ```

2. 在 Python 腳本中使用：
   ```python
   import clr
   import sys
   import os

   # 添加 RdpClientBridge.dll 的路徑
   dll_path = os.path.join('path', 'to', 'RdpClientBridge.dll')
   clr.AddReference(dll_path)

   # 導入命名空間
   from RdpClientBridge import RDPConnection

   # 創建 RDP 連線實例
   rdp_conn = RDPConnection()
   rdp_conn.Server = "127.0.0.1"
   rdp_conn.Username = "test_user"
   rdp_conn.Password = "test_password"
   rdp_conn.Port = 3389
   rdp_conn.Width = 1024
   rdp_conn.Height = 768
   rdp_conn.ColorDepth = 16
   rdp_conn.Fullscreen = False

   # 開始連線
   rdp_conn.Connect()
   ```

#### 使用 COM 介面

1. 以系統管理員權限執行命令提示字元
2. 註冊 DLL： ()
   ```cmd
   regasm.exe /tlb /codebase "path\to\RdpClientBridge.dll"
   ```

3. 在 Python 中使用：
   ```python
   import comtypes.client

   # 使用 GUID 實例化對象
   CLSID = "{7F4C6F4B-C7A7-4859-8A11-923A3E8A5714}" 
   rdp_client = comtypes.client.CreateObject(CLSID)

   # 設定連線參數
   rdp_client.Server = "127.0.0.1"
   rdp_client.Username = "test_user"
   rdp_client.Password = "test_password"
   rdp_client.Port = 3389
   rdp_client.Width = 1024
   rdp_client.Height = 768

   # 開始連線
   rdp_client.Connect()
   ```

## 使用方法

### C# 中的使用

```csharp
using RdpClientBridge;

// 創建 RDP 連線實例
RDPConnection rdpConn = new RDPConnection("127.0.0.1", "username", "password", 3389, 1024, 768, 16, false);

// 或使用預設建構函數
RDPConnection rdpConn = new RDPConnection();
rdpConn.Server = "127.0.0.1";
rdpConn.Username = "username";
rdpConn.Password = "password";
rdpConn.Port = 3389;
rdpConn.Width = 1024;
rdpConn.Height = 768;
rdpConn.ColorDepth = 16;
rdpConn.Fullscreen = false;

// 開始連線
rdpConn.Connect();

// 將視窗移至後台（可選，但仍保持連線）
rdpConn.MoveToBackground();

// 在後台發送按鍵 (例如發送 'A' 鍵，虛擬鍵碼 65)
rdpConn.SendKeyBackground(65);

// 在後台發送滑鼠點擊 (座標 100, 200)
rdpConn.SendMouseClickBackground(100, 200);

// 恢復視窗至前景（如果需要）
// rdpConn.RestoreWindow();
```

### Python 中的使用

參考 `python_example/multi_rdp_manager.py` 中的完整範例，該範例展示了如何在 Python 中安全地創建 STA 執行緒來處理 RDP 連線。

專案支援後台鍵鼠操作，無需將 RDP 視窗置於前景即可發送鍵盤和滑鼠輸入。範例檔案 `python_example/multi_rdp_manager.py` 包含了以下自動化 API：

- `click(x, y)`: 在指定座標點擊滑鼠
- `right_click(x, y)`: 在指定座標點擊滑鼠右鍵
- `mouse_down(x, y)`: 在指定座標按下滑鼠
- `mouse_up(x, y)`: 在指定座標放開滑鼠
- `mouse_move(x, y, is_left_down=False)`: 移動滑鼠（支援拖曳狀態）
- `press_key(key_code_or_char)`: 發送按鍵 (支援按鍵名稱如 'ENTER', 'A', 'LWIN' 或整數鍵碼)
- `key_down(key_code_or_char)`: 按下按鍵 (支援按鍵名稱或整數鍵碼)
- `key_up(key_code_or_char)`: 放開按鍵 (支援按鍵名稱或整數鍵碼)
- `type_text(text, interval=0.05)`: 輸入字串
- `hide_window()`: 隱藏視窗 (背景模式)
- `show_window()`: 顯示視窗

## API 參考

### 屬性
- `Server` (string): RDP 伺服器地址
- `Username` (string): 使用者名稱
- `Password` (string): 密碼
- `Port` (int): RDP 連接埠 (預設 3389)
- `Width` (int): 桌面寬度 (預設 1024)
- `Height` (int): 桌面高度 (預設 768)
- `ColorDepth` (int): 顏色深度 (預設 16)
- `Fullscreen` (bool): 是否全螢幕 (預設 False)

### 方法
- `Connect()`: 開始 RDP 連接
- `Disconnect()`: 斷開 RDP 連接
- `MoveToBackground()`: 將 RDP 視窗移至後台（視窗隱藏但仍保持連線）
- `RestoreWindow()`: 恢復 RDP 視窗至前景
- `SendKeyBackground(int virtualKeyCode)`: 在後台發送按鍵（無需視窗焦點）
- `SendKeyDown(int virtualKeyCode)`: 在後台發送按鍵按下（無需視窗焦點）
- `SendKeyUp(int virtualKeyCode)`: 在後台發送按鍵放開（無需視窗焦點）
- `SendMouseClickBackground(int x, int y)`: 在後台發送滑鼠點擊（無需視窗焦點）
- `SendMouseRightClickBackground(int x, int y)`: 在後台發送滑鼠右鍵點擊（無需視窗焦點）
- `SendMouseDownBackground(int x, int y)`: 在後台發送滑鼠按下（無需視窗焦點）
- `SendMouseUpBackground(int x, int y)`: 在後台發送滑鼠放開（無需視窗焦點）
- `SendMouseMoveBackground(int x, int y, bool isLeftDown)`: 在後台發送滑鼠移動（支援拖曳狀態）

## 重要注意事項

1. **權限要求**：註冊 DLL 需要管理員權限
2. **相依性**：確保目標系統安裝了 .NET Framework 4.8 或更高版本
3. **安全性**：在生產環境中，不要在代碼中硬編碼用戶名和密碼
4. **錯誤處理**：始終包含適當的錯誤處理邏輯
5. **線程處理**：RDP 連接會啟動 UI 線程，注意線程安全

## 檔案結構

```
RdpClientBridge/
├── RdpClientBridge.sln          # Visual Studio 解決方案
├── RdpClientBridge/             # C# 主專案目錄
│   ├── RdpClientBridge.csproj   # C# 專案文件
│   ├── Program.cs              # 主程式入口點
│   ├── RDPConnection.cs        # RDP 連線控制類別
│   ├── RDPManager.cs           # RDP 管理器類別
│   └── Properties/             # 專案屬性
├── python_example/             # Python 整合範例
│   ├── rdp_connector.py        # Python 連接器範例
│   ├── RdpClientBridge.dll     # 已編譯的組件
│   ├── Interop.MSTSCLib.dll    # COM 互操作組件
│   └── AxInterop.MSTSCLib.dll  # ActiveX 互操作組件
├── PYTHON_INTEGRATION.md       # Python 整合詳細指南
└── README.md                   # 本文件
```

## 授權

此專案採用 [MIT 授權](LICENSE)（如 LICENSE 文件存在）。