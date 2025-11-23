# Python 整合指南

本指南詳細說明如何從 Python 調用 RDP Client Bridge，提供兩種不同的方法：

1. 使用 COM 和 comtypes
2. 使用 pythonnet

## 方法一：使用 COM 和 comtypes

### 步驟 1：註冊 DLL

在使用 COM 介面之前，您需要註冊 RdpClientBridge.dll：

```cmd
# 使用管理員權限的命令提示字元
regasm.exe /tlb /codebase "path\to\RdpClientBridge.dll"
```

### 步驟 2：Python 調用範例

```python
import comtypes.client
import time

# 使用您在 C# 類別上定義的 GUID 或 ProgID 來實例化對象
# 使用 ProgID (如果沒有自訂 ProgID，它通常是 Namespace.Class):
# ProgID = "RdpClientBridge.RDPConnection" 
# 或使用您在 RDPConnection 類別中定義的 GUID:
CLSID = "{7F4C6F4B-C7A7-4859-8A11-923A3E8A5714}" 

try:
    # 創建 COM 對象
    rdp_client = comtypes.client.CreateObject(CLSID)

    # 設定連線參數（通過我們在 C# 中定義的屬性）
    rdp_client.Server = "127.0.0.1"
    rdp_client.Username = "test_user"
    rdp_client.Password = "test_password"
    rdp_client.Port = 3389
    rdp_client.Width = 1024
    rdp_client.Height = 768

    print("RDP Client object created and parameters set.")
    
    # 調用 Connect 方法
    rdp_client.Connect()

    # 由於 Connect() 會啟動一個阻塞的 Windows 視窗，
    # Python 腳本會在這裡等待 RDP 視窗關閉。

except Exception as e:
    print(f"Error calling RDP Client: {e}")
```

## 方法二：使用 pythonnet

### 步驟 1：安裝 pythonnet

```bash
pip install pythonnet
```

### 步驟 2：Python 調用範例

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

## 完整的高級範例

參見 `advanced_python_example.py` 文件，它包含了：

- 自動 DLL 註冊功能
- 兩種調用方法的實現
- 錯誤處理
- 交互式選擇調用方式

## 重要注意事項

1. **權限要求**：註冊 DLL 需要管理員權限
2. **相依性**：確保目標系統安裝了 .NET Framework 4.8 或更高版本
3. **安全性**：在生產環境中，不要在代碼中硬編碼用戶名和密碼
4. **錯誤處理**：始終包含適當的錯誤處理邏輯
5. **線程處理**：RDP 連接會啟動 UI 線程，注意線程安全

## 故障排除

### COM 註冊問題
- 確保以管理員身份運行命令提示字元
- 確保安裝了 .NET Framework SDK
- 檢查 DLL 路徑是否正確

### pythonnet 問題
- 確保安裝了正確版本的 pythonnet
- 確保 Python 是 64 位版本（與 DLL 架構匹配）

## 可用屬性和方法

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