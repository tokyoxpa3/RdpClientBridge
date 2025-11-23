using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

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
        int ColorDepth { get; set; }
        bool Fullscreen { get; set; }
    }

    [ComVisible(true)]
    [Guid("7F4C6F4B-C7A7-4859-8A11-923A3E8A5714")]
    [ClassInterface(ClassInterfaceType.None)]
    public class RDPConnection : Form, IRDPConnection
    {
        private AxMSTSCLib.AxMsRdpClient2 rdpControl;
        
        public string Server { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public int Port { get; set; }
        public new int Width { get; set; } // 'new' to hide Form.Width
        public new int Height { get; set; } // 'new' to hide Form.Height
        public int ColorDepth { get; set; }
        public bool Fullscreen { get; set; }

        public RDPConnection()
        {
            // Set default values
            this.Port = 3389;
            this.Width = 1024;
            this.Height = 768;
            this.ColorDepth = 16;
            this.Fullscreen = false;

            // Initialize components and events
            InitializeComponent();
            InitializeControlEvents();
        }
        
        private void InitializeComponent()
        {
            this.rdpControl = new AxMSTSCLib.AxMsRdpClient2();
            
            ((System.ComponentModel.ISupportInitialize)(this.rdpControl)).BeginInit();
            this.SuspendLayout();

            // Setup the Form
            this.ClientSize = new System.Drawing.Size(this.Width, this.Height);
            this.Name = "RDPConnectionForm";
            this.StartPosition = FormStartPosition.CenterScreen;

            // Setup the RDP control
            this.Controls.Add(this.rdpControl);
            this.rdpControl.Dock = DockStyle.Fill;
            this.rdpControl.Enabled = true;
            this.rdpControl.Name = "rdpControl";
            
            ((System.ComponentModel.ISupportInitialize)(this.rdpControl)).EndInit();
            this.ResumeLayout(false);
        }

        private void InitializeControlEvents()
        {
            this.Load += (sender, e) => {
                try
                {
                    this.rdpControl.Connect();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"RDP Connection Error on Connect: {ex.Message}", "RDP Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    this.Close();
                }
            };
            
            this.FormClosing += (sender, e) => {
                this.Disconnect();
            };
            
            rdpControl.OnConnecting += (s, e) => Console.WriteLine("RDP: Connecting...");
            rdpControl.OnConnected += (s, e) => Console.WriteLine("RDP: Connected successfully.");
            rdpControl.OnLoginComplete += (s, e) => Console.WriteLine("RDP: Login complete.");
            rdpControl.OnDisconnected += (s, e) => Console.WriteLine($"RDP: Disconnected. Reason: {e.discReason}");
        }

        public void Disconnect()
        {
            if (this.rdpControl != null && this.rdpControl.Connected != 0)
            {
                this.rdpControl.Disconnect();
            }
        }

        // This method now *applies* settings before the form is shown
        public void Connect()
        {
            // Apply settings to the control
            this.rdpControl.Server = this.Server;
            this.rdpControl.UserName = this.Username;
            this.rdpControl.AdvancedSettings2.ClearTextPassword = this.Password;
            this.rdpControl.ColorDepth = this.ColorDepth;
            this.rdpControl.DesktopWidth = this.Width;
            this.rdpControl.DesktopHeight = this.Height;
            this.rdpControl.FullScreen = this.Fullscreen;
            
            // Advanced Settings
            this.rdpControl.AdvancedSettings2.AcceleratorPassthrough = -1;
            this.rdpControl.AdvancedSettings2.Compress = -1;
            this.rdpControl.AdvancedSettings2.BitmapPersistence = -1;
            this.rdpControl.AdvancedSettings2.CachePersistenceActive = -1;
            
            // Redirection Settings
            this.rdpControl.AdvancedSettings2.RedirectDrives = false;
            this.rdpControl.AdvancedSettings2.RedirectPrinters = false;
            this.rdpControl.AdvancedSettings2.RedirectPorts = false;
            this.rdpControl.AdvancedSettings2.RedirectSmartCards = false;
            
            if (this.Port != 3389)
            {
                this.rdpControl.AdvancedSettings2.RDPPort = this.Port;
            }

            // Update form properties
            this.Text = $"RDP - {this.Username}@{this.Server}:{this.Port}";
            this.Size = new Size(this.Width, this.Height);
        }
    }
}