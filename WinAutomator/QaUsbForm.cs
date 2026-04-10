using System;
using System.Drawing;
using System.Windows.Forms;

namespace WinAutomator
{
    public class QaUsbForm : Form
    {
        private Label lblStatus;
        private Button btnSkip;
        
        public QaUsbForm()
        {
            this.Text = "Hardware Check: USB Ports (מבדק כניסות USB)";
            this.Size = new Size(500, 300);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.FromArgb(40, 40, 40);
            this.RightToLeft = RightToLeft.Yes;
            
            Label lblInstr = new Label() { 
                Text = "מבדק שקעי USB", 
                Dock = DockStyle.Top, Height = 60, ForeColor = Color.Cyan, 
                Font = new Font("Segoe UI", 16, FontStyle.Bold), TextAlign = ContentAlignment.MiddleCenter 
            };
            this.Controls.Add(lblInstr);

            lblStatus = new Label() { 
                Text = "אנא הכנס התקן שמע, עכבר או דיסק-און-קי לאחד מהשקעים במחשב...\nאני מאזין למערכת ההפעלה כעת...", 
                Location = new Point(50, 90), Size = new Size(400, 80), ForeColor = Color.White, 
                Font = new Font("Segoe UI", 12), TextAlign = ContentAlignment.MiddleCenter 
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
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            // WM_DEVICECHANGE (0x0219)
            if (m.Msg == 0x0219)
            {
                // DBT_DEVICEARRIVAL (0x8000)
                if (m.WParam.ToInt32() == 0x8000)
                {
                    UsbConnected();
                }
            }
        }

        private void UsbConnected()
        {
            if (InvokeRequired) { Invoke(new Action(UsbConnected)); return; }

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
