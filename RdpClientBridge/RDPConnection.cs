using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.Text;
using System.Collections.Generic;

namespace RdpClientBridge
{
    // --- 介面宣告 (保持不變) ---
    [ComVisible(true)]
    [Guid("18D9D99F-65F9-46B4-A3D9-92F6786A0108")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IRDPConnection
    {
        void Connect();
        void Disconnect();
        string Server { get; set; }
        string Username { get; set; }
        string Password { get; set; }
        int Port { get; set; }
        int Width { get; set; }
        int Height { get; set; }
        bool Fullscreen { get; set; }
        void MoveToBackground();
        void RestoreWindow();
        void SendKeyBackground(int virtualKeyCode);
        void SendMouseClickBackground(int x, int y);
    }

    [ComVisible(true)]
    [Guid("7F4C6F4B-C7A7-4859-8A11-923A3E8A5714")]
    [ClassInterface(ClassInterfaceType.None)]
    public class RDPConnection : Form, IRDPConnection
    {
        private AxMSTSCLib.AxMsRdpClient10 rdpControl;

        // --- P/Invoke ---
        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern uint MapVirtualKey(uint uCode, uint uMapType);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr hwndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        // Windows Constants
        private const uint WM_KEYDOWN = 0x0100;
        private const uint WM_KEYUP = 0x0101;
        private const uint WM_LBUTTONDOWN = 0x0201;
        private const uint WM_LBUTTONUP = 0x0202;
        private const uint WM_MOUSEMOVE = 0x0200;
        private const uint WM_SETFOCUS = 0x0007; // 關鍵：強制設定焦點
        private const uint MK_LBUTTON = 0x0001;
        private const uint MAPVK_VK_TO_VSC = 0x00;

        // Properties
        public string Server { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public int Port { get; set; } = 3389;
        public new int Width { get; set; } = 1024;
        public new int Height { get; set; } = 768;
        public bool Fullscreen { get; set; } = false;

        public RDPConnection()
        {
            InitializeComponent();
            InitializeControlEvents();
        }

        private void InitializeComponent()
        {
            this.rdpControl = new AxMSTSCLib.AxMsRdpClient10();
            ((System.ComponentModel.ISupportInitialize)(this.rdpControl)).BeginInit();
            this.SuspendLayout();
            this.ClientSize = new Size(this.Width, this.Height);
            this.Controls.Add(this.rdpControl);
            this.rdpControl.Dock = DockStyle.Fill;
            this.rdpControl.Enabled = true;
            ((System.ComponentModel.ISupportInitialize)(this.rdpControl)).EndInit();
            this.ResumeLayout(false);
        }

        private void InitializeControlEvents()
        {
            this.Load += (s, e) => { try { this.rdpControl.Connect(); } catch { } };
            this.FormClosing += (s, e) => Disconnect();
            rdpControl.OnConnected += (s, e) => Console.WriteLine("C# Log: RDP Connected.");
        }

        public void Connect()
        {
            this.rdpControl.Server = this.Server;
            this.rdpControl.UserName = this.Username;
            this.rdpControl.AdvancedSettings2.ClearTextPassword = this.Password;
            this.rdpControl.DesktopWidth = this.Width;
            this.rdpControl.DesktopHeight = this.Height;
            this.rdpControl.AdvancedSettings2.RDPPort = this.Port;
            this.rdpControl.AdvancedSettings2.SmartSizing = true;
            this.Text = $"RDP - {this.Username}@{this.Server}";
            this.Size = new Size(this.Width, this.Height);
        }

        public void Disconnect() { if (this.rdpControl.Connected != 0) this.rdpControl.Disconnect(); }

        public void MoveToBackground()
        {
            this.Invoke(new MethodInvoker(() => { this.Location = new Point(-3000, -3000); }));
        }

        public void RestoreWindow()
        {
            this.Invoke(new MethodInvoker(() => { this.Location = new Point(100, 100); this.BringToFront(); }));
        }

        // --- 核心改進部分 ---

        // 遞迴尋找特定的類別名稱
        private IntPtr FindChildWindowByClass(IntPtr parent, string className)
        {
            IntPtr result = IntPtr.Zero;
            EnumChildWindows(parent, (hwnd, param) => {
                StringBuilder sb = new StringBuilder(256);
                GetClassName(hwnd, sb, sb.Capacity);
                if (sb.ToString() == className)
                {
                    result = hwnd;
                    return false; // Stop enumeration
                }
                return true; // Continue
            }, IntPtr.Zero);
            return result;
        }

        private IntPtr GetInputHandlerWindow()
        {
            if (this.rdpControl.Handle == IntPtr.Zero) return IntPtr.Zero;

            // 策略 1: 標準路徑 (Form -> UIContainer -> IHWindow)
            // 先找 UIContainerClass
            IntPtr uiContainer = FindChildWindowByClass(this.rdpControl.Handle, "UIContainerClass");
            if (uiContainer != IntPtr.Zero)
            {
                // 在 Container 裡面找 IHWindowClass
                IntPtr ihWindow = FindChildWindowByClass(uiContainer, "IHWindowClass");
                if (ihWindow != IntPtr.Zero)
                {
                    // Console.WriteLine($"Debug: Found IHWindowClass (HWND: {ihWindow})");
                    return ihWindow;
                }

                // 如果找不到 IHWindow，有些版本直接用 UIContainer 接收輸入
                // Console.WriteLine($"Debug: IHWindow not found, using UIContainer (HWND: {uiContainer})");
                return uiContainer;
            }

            // 策略 2: 如果找不到結構，回傳主控制項 (雖然通常無效，但死馬當活馬醫)
            Console.WriteLine("Debug: Could not find internal RDP windows. Using Main Handle.");
            return this.rdpControl.Handle;
        }

        private IntPtr MakeKeyLParam(uint scanCode, bool isDown)
        {
            int lParam = 1; // Repeat count
            lParam |= ((int)scanCode << 16);
            if (!isDown)
            {
                lParam |= (1 << 30); // Previous state
                lParam |= (1 << 31); // Transition
            }
            return (IntPtr)lParam;
        }

        public void SendKeyBackground(int virtualKeyCode)
        {
            IntPtr targetHwnd = GetInputHandlerWindow();
            if (targetHwnd == IntPtr.Zero) return;

            uint scanCode = MapVirtualKey((uint)virtualKeyCode, MAPVK_VK_TO_VSC);

            // Debug: 檢查 ScanCode 是否正確 (如果為 0，RDP 絕對沒反應)
            Console.WriteLine($"C# Log: Sending VK={virtualKeyCode} -> ScanCode={scanCode} to HWND={targetHwnd}");

            // 關鍵修正：發送 WM_SETFOCUS。
            // 這是 "騙" RDP 視窗它現在是被選中的，否則它會忽略輸入。
            SendMessage(targetHwnd, WM_SETFOCUS, IntPtr.Zero, IntPtr.Zero);

            // 發送按鍵
            PostMessage(targetHwnd, WM_KEYDOWN, (IntPtr)virtualKeyCode, MakeKeyLParam(scanCode, true));
            // 稍微等待一點點模擬真實按鍵長度 (雖非必須，但建議)
            Thread.Sleep(20);
            PostMessage(targetHwnd, WM_KEYUP, (IntPtr)virtualKeyCode, MakeKeyLParam(scanCode, false));
        }

        public void SendMouseClickBackground(int x, int y)
        {
            IntPtr targetHwnd = GetInputHandlerWindow();
            if (targetHwnd == IntPtr.Zero) return;

            // 這裡也發送 Focus，確保滑鼠點擊被處理
            SendMessage(targetHwnd, WM_SETFOCUS, IntPtr.Zero, IntPtr.Zero);

            IntPtr lParam = (IntPtr)((y << 16) | (x & 0xFFFF));
            PostMessage(targetHwnd, WM_LBUTTONDOWN, (IntPtr)MK_LBUTTON, lParam);
            PostMessage(targetHwnd, WM_LBUTTONUP, IntPtr.Zero, lParam);

            Console.WriteLine($"C# Log: Click ({x},{y}) sent to HWND={targetHwnd}");
        }
    }
}