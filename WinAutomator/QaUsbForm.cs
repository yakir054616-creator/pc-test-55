using System;
using System.Drawing;
using System.Management;
using System.Windows.Forms;

namespace WinAutomator
{
    public class QaUsbForm : Form
    {
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        private Label lblStatus;
        private Button btnSkip;
        private bool _detected = false; // prevent double-fire
        private int _initialUsbCount = 0;

        public QaUsbForm()
        {
            this.Text = "Project Aura - USB Check";
            this.Size = new Size(500, 300);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.None;
            this.MaximizeBox = false;
            this.BackColor = Color.FromArgb(28, 28, 28);
            this.RightToLeft = RightToLeft.No;

            this.Paint += (s, e) => {
                using (Pen p = new Pen(Color.FromArgb(70, 70, 70), 2))
                    e.Graphics.DrawRectangle(p, 0, 0, this.Width - 1, this.Height - 1);
            };

            Panel titleBar = new Panel() { Dock = DockStyle.Top, Height = 40, BackColor = Color.FromArgb(20, 20, 20), Cursor = Cursors.Default };
            titleBar.MouseDown += (s, e) => {
                if (e.Button == MouseButtons.Left) { ReleaseCapture(); SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0); }
            };
            Label lblTitle = new Label() { Text = "מבדק שקעי USB", ForeColor = Color.Cyan, Font = new Font("Segoe UI", 12, FontStyle.Bold), Location = new Point(10, 8), AutoSize = true, BackColor = Color.Transparent };
            titleBar.Controls.Add(lblTitle);

            Label lblClose = new Label() { Text = "✕", Font = new Font("Segoe UI", 12, FontStyle.Bold), Location = new Point(this.Width - 30, 8), AutoSize = true, ForeColor = Color.DarkGray, Cursor = Cursors.Hand, BackColor = Color.Transparent };
            lblClose.MouseEnter += (s, e) => lblClose.ForeColor = Color.LightCoral;
            lblClose.MouseLeave += (s, e) => lblClose.ForeColor = Color.DarkGray;
            lblClose.Click += (s, e) => { this.DialogResult = DialogResult.Ignore; this.Close(); };
            titleBar.Controls.Add(lblClose);
            this.Controls.Add(titleBar);

            lblStatus = new Label() {
                Text = "אנא הכנס התקן שמע, עכבר או דיסק-און-קי לאחד מהשקעים במחשב...\nאני מאזין למערכת ההפעלה כעת...",
                Location = new Point(50, 90), Size = new Size(400, 80), ForeColor = Color.White,
                Font = new Font("Segoe UI", 12), TextAlign = ContentAlignment.MiddleCenter,
                RightToLeft = RightToLeft.Yes
            };
            this.Controls.Add(lblStatus);

            btnSkip = new Button() {
                Text = "דלג על בדיקה זו", Location = new Point(175, 200), Size = new Size(150, 40),
                BackColor = Color.FromArgb(100, 100, 100), ForeColor = Color.White,
                Font = new Font("Segoe UI", 10, FontStyle.Bold), Cursor = Cursors.Hand
            };
            btnSkip.FlatAppearance.BorderSize = 0;
            btnSkip.FlatStyle = FlatStyle.Flat;
            btnSkip.Click += (s, e) => { this.DialogResult = DialogResult.Ignore; this.Close(); };
            this.Controls.Add(btnSkip);

            this.Shown += (s, e) => {
                _initialUsbCount = GetUsbDeviceCount();
            };
        }

        // ── Primary detection: WndProc catches WM_DEVICECHANGE ──
        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == 0x0219) // WM_DEVICECHANGE
            {
                int wp = m.WParam.ToInt32();
                // 0x8000 = DBT_DEVICEARRIVAL  (storage, audio…)
                // 0x0007 = DBT_DEVNODES_CHANGED (any hardware changed)
                if (wp == 0x8000 || wp == 0x0007)
                {
                    System.Threading.Tasks.Task.Run(() =>
                    {
                        int currentCount = GetUsbDeviceCount();
                        if (currentCount > _initialUsbCount)
                        {
                            if (IsHandleCreated && !_detected)
                            {
                                this.BeginInvoke(new Action(UsbConnected));
                            }
                        }
                        // Always update baseline so unplugging and re-plugging works
                        _initialUsbCount = currentCount;
                    });
                }
            }
        }

        private int GetUsbDeviceCount()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE PNPDeviceID LIKE 'USB%'"))
                {
                    return searcher.Get().Count;
                }
            }
            catch { return 0; }
        }

        private void UsbConnected()
        {
            if (_detected) return; // prevent double-fire from both methods
            _detected = true;

            lblStatus.Text = "זוהה התקן USB שחובר בהצלחה! ✓";
            lblStatus.ForeColor = Color.LimeGreen;
            btnSkip.Visible = false;

            System.Windows.Forms.Timer t = new System.Windows.Forms.Timer() { Interval = 1500 };
            t.Tick += (s, e) => {
                t.Stop();
                t.Dispose();
                this.DialogResult = DialogResult.Yes;
                this.Close();
            };
            t.Start();
        }
    }
}
