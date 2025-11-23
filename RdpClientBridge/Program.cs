using System;
using System.Windows.Forms;

namespace RdpClientBridge
{
    public class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            // 測試 RDP 連線功能
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            
            // 使用範例參數進行測試
            // 注意：在實際使用時，需要替換為有效的 RDP 伺服器資訊
            try
            {
                RDPConnection rdpConn = new RDPConnection("127.0.0.1", "username", "password", 3389, 1024, 768, 16, false);
                rdpConn.Connect();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "RDP Client Bridge", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}