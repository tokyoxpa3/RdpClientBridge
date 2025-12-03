using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Threading;
using System.Text;

namespace RdpClientBridge
{
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

        // --- 鍵盤控制 ---
        void SendKeyBackground(int virtualKeyCode);
        // [新增] 獨立的按下與放開
        void SendKeyDown(int virtualKeyCode);
        void SendKeyUp(int virtualKeyCode);

        // --- 滑鼠控制 ---
        void SendMouseClickBackground(int x, int y);
        void SendMouseRightClickBackground(int x, int y);
        void SendMouseDownBackground(int x, int y);
        void SendMouseUpBackground(int x, int y);
        void SendMouseMoveBackground(int x, int y, bool isLeftDown);
    }

    [ComVisible(true)]
    [Guid("7F4C6F4B-C7A7-4859-8A11-923A3E8A5714")]
    [ClassInterface(ClassInterfaceType.None)]
    public class RDPConnection : Form, IRDPConnection
    {
        private AxMSTSCLib.AxMsRdpClient10 rdpControl;

        // --- P/Invoke (保持不變) ---
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
        private const uint WM_RBUTTONDOWN = 0x0204;
        private const uint WM_RBUTTONUP = 0x0205;
        private const uint WM_MOUSEMOVE = 0x0200;
        private const uint WM_SETFOCUS = 0x0007;
        private const uint MK_LBUTTON = 0x0001;
        private const uint MK_RBUTTON = 0x0002;
        private const uint MAPVK_VK_TO_VSC = 0x00;

        // Properties (保持不變)
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

        // --- 核心邏輯 (完全保留您寫好的) ---

        private IntPtr FindChildWindowByClass(IntPtr parent, string className)
        {
            IntPtr result = IntPtr.Zero;
            EnumChildWindows(parent, (hwnd, param) => {
                StringBuilder sb = new StringBuilder(256);
                GetClassName(hwnd, sb, sb.Capacity);
                if (sb.ToString() == className)
                {
                    result = hwnd;
                    return false;
                }
                return true;
            }, IntPtr.Zero);
            return result;
        }

        private IntPtr GetInputHandlerWindow()
        {
            if (this.rdpControl.Handle == IntPtr.Zero) return IntPtr.Zero;

            // 策略 1
            IntPtr uiContainer = FindChildWindowByClass(this.rdpControl.Handle, "UIContainerClass");
            if (uiContainer != IntPtr.Zero)
            {
                IntPtr ihWindow = FindChildWindowByClass(uiContainer, "IHWindowClass");
                if (ihWindow != IntPtr.Zero) return ihWindow;
                return uiContainer;
            }
            // 策略 2
            return this.rdpControl.Handle;
        }

        private IntPtr MakeKeyLParam(uint scanCode, bool isDown)
        {
            int lParam = 1;
            lParam |= ((int)scanCode << 16);
            if (!isDown)
            {
                lParam |= (1 << 30);
                lParam |= (1 << 31);
            }
            return (IntPtr)lParam;
        }

        // --- [修改部分] 拆分 KeyDown 和 KeyUp ---

        public void SendKeyDown(int virtualKeyCode)
        {
            IntPtr targetHwnd = GetInputHandlerWindow();
            if (targetHwnd == IntPtr.Zero) return;

            uint scanCode = MapVirtualKey((uint)virtualKeyCode, MAPVK_VK_TO_VSC);

            // 保持您的邏輯：先搶焦點
            SendMessage(targetHwnd, WM_SETFOCUS, IntPtr.Zero, IntPtr.Zero);

            // 發送 KeyDown
            // 注意 MakeKeyLParam 的第二個參數是 true (isDown)
            PostMessage(targetHwnd, WM_KEYDOWN, (IntPtr)virtualKeyCode, MakeKeyLParam(scanCode, true));
        }

        public void SendKeyUp(int virtualKeyCode)
        {
            IntPtr targetHwnd = GetInputHandlerWindow();
            if (targetHwnd == IntPtr.Zero) return;

            uint scanCode = MapVirtualKey((uint)virtualKeyCode, MAPVK_VK_TO_VSC);

            // 保持您的邏輯：先搶焦點 (確保 KeyUp 也被接收)
            SendMessage(targetHwnd, WM_SETFOCUS, IntPtr.Zero, IntPtr.Zero);

            // 發送 KeyUp
            // 注意 MakeKeyLParam 的第二個參數是 false (!isDown)
            PostMessage(targetHwnd, WM_KEYUP, (IntPtr)virtualKeyCode, MakeKeyLParam(scanCode, false));
        }

        // 為了相容性，原本的 SendKeyBackground 直接呼叫上面兩個方法
        public void SendKeyBackground(int virtualKeyCode)
        {
            SendKeyDown(virtualKeyCode);
            Thread.Sleep(20);
            SendKeyUp(virtualKeyCode);
        }

        // --- 滑鼠部分保持不變 ---
        public void SendMouseClickBackground(int x, int y)
        {
            IntPtr targetHwnd = GetInputHandlerWindow();
            if (targetHwnd == IntPtr.Zero) return;
            SendMessage(targetHwnd, WM_SETFOCUS, IntPtr.Zero, IntPtr.Zero);
            IntPtr lParam = (IntPtr)((y << 16) | (x & 0xFFFF));
            PostMessage(targetHwnd, WM_LBUTTONDOWN, (IntPtr)MK_LBUTTON, lParam);
            PostMessage(targetHwnd, WM_LBUTTONUP, IntPtr.Zero, lParam);
        }

        public void SendMouseRightClickBackground(int x, int y)
        {
            IntPtr targetHwnd = GetInputHandlerWindow();
            if (targetHwnd == IntPtr.Zero) return;
            SendMessage(targetHwnd, WM_SETFOCUS, IntPtr.Zero, IntPtr.Zero);
            IntPtr lParam = (IntPtr)((y << 16) | (x & 0xFFFF));
            PostMessage(targetHwnd, WM_RBUTTONDOWN, (IntPtr)MK_RBUTTON, lParam);
            PostMessage(targetHwnd, WM_RBUTTONUP, IntPtr.Zero, lParam);
        }

        public void SendMouseDownBackground(int x, int y)
        {
            IntPtr targetHwnd = GetInputHandlerWindow();
            if (targetHwnd == IntPtr.Zero) return;
            SendMessage(targetHwnd, WM_SETFOCUS, IntPtr.Zero, IntPtr.Zero);
            IntPtr lParam = (IntPtr)((y << 16) | (x & 0xFFFF));
            PostMessage(targetHwnd, WM_LBUTTONDOWN, (IntPtr)MK_LBUTTON, lParam);
        }

        public void SendMouseUpBackground(int x, int y)
        {
            IntPtr targetHwnd = GetInputHandlerWindow();
            if (targetHwnd == IntPtr.Zero) return;
            IntPtr lParam = (IntPtr)((y << 16) | (x & 0xFFFF));
            PostMessage(targetHwnd, WM_LBUTTONUP, IntPtr.Zero, lParam);
        }

        public void SendMouseMoveBackground(int x, int y, bool isLeftDown)
        {
            IntPtr targetHwnd = GetInputHandlerWindow();
            if (targetHwnd == IntPtr.Zero) return;
            IntPtr lParam = (IntPtr)((y << 16) | (x & 0xFFFF));
            IntPtr wParam = isLeftDown ? (IntPtr)MK_LBUTTON : IntPtr.Zero;
            PostMessage(targetHwnd, WM_MOUSEMOVE, wParam, lParam);
        }
    }
}