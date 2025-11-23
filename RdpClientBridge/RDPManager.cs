using System;
using System.Drawing;
using System.Windows.Forms;
using AxMSTSCLib;
using MSTSCLib;

namespace RdpClientBridge
{
    public static class RDPManager
    {
        public static void StartRdpSession(string server, string user, string password, int port = 3389, int width = 1024, int height = 768, int colorDepth = 16, bool fullscreen = false)
        {
            // 確保您的 DLL 可以存取 mstscax.dll
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 創建 RDP 控制項和視窗
            Form form = new Form();
            AxMsRdpClient9NotSafeForScripting rdp = new AxMsRdpClient9NotSafeForScripting();

            try
            {
                // 配置 RDP 控制項
                rdp.Server = server;
                rdp.UserName = user;
                rdp.AdvancedSettings2.ClearTextPassword = password;

                // 設定連接參數 
                rdp.ColorDepth = colorDepth; // 顏色深度
                rdp.DesktopWidth = width;
                rdp.DesktopHeight = height;
                rdp.FullScreen = fullscreen;

                // 設定進階參數 
                rdp.AdvancedSettings2.AcceleratorPassthrough = -1; // 啟用加速鍵傳遞
                rdp.AdvancedSettings2.Compress = -1; // 啟用壓縮
                rdp.AdvancedSettings2.BitmapPersistence = -1; // 啟用位圖持久性
                rdp.AdvancedSettings2.BitmapPeristence = -1; // 另一個位圖持久性設置
                rdp.AdvancedSettings2.CachePersistenceActive = -1; // 啟用快取持久性
                
                // 設定自訂端口
                if (port != 3389)
                {
                    rdp.AdvancedSettings2.RDPPort = port;
                }

                // 連接事件處理 
                rdp.OnConnecting += Rdp_OnConnecting;
                rdp.OnConnected += Rdp_OnConnected;
                rdp.OnLoginComplete += Rdp_OnLoginComplete;
                rdp.OnDisconnected += Rdp_OnDisconnected;

                // 設定視窗和控制項佈局
                form.Controls.Add(rdp);
                rdp.Dock = DockStyle.Fill;
                form.Text = $"RDP Client - {user}@{server}:{port}";
                form.Size = new Size(width + 40, height + 60); // 稍微大一點以容納標題列和狀態列
                form.StartPosition = FormStartPosition.CenterScreen;
                
                // 當表單關閉時，斷開連接並釋放資源
                form.FormClosing += (sender, e) => {
                    if (rdp.Connected != 0)
                    {
                        rdp.Disconnect();
                    }
                    rdp.Dispose();
                };

                // 連線
                rdp.Connect();

                // 運行視窗
                Application.Run(form);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"RDP Connection Error: {ex.Message}", "RDP Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
                // 清理資源
                rdp?.Dispose();
                form?.Dispose();
            }
        }

        // 事件處理方法
        private static void Rdp_OnConnecting(object sender, EventArgs e)
        {
            Console.WriteLine("RDP: Connecting...");
        }

        private static void Rdp_OnConnected(object sender, EventArgs e)
        {
            Console.WriteLine("RDP: Connected successfully.");
        }

        private static void Rdp_OnLoginComplete(object sender, EventArgs e)
        {
            Console.WriteLine("RDP: Login complete.");
        }

        private static void Rdp_OnDisconnected(object sender, IMsTscAxEvents_OnDisconnectedEvent e)
        {
            Console.WriteLine($"RDP: Disconnected. Reason: {e.discReason}");
        }
    }
}
