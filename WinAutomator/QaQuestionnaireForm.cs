using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;

namespace WinAutomator
{
    public class QaQuestionnaireForm : Form
    {
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        // Outputs
        public string AnsPlastic { get; set; } = "כן, המחשב שלם";
        public string AnsOldDisk { get; set; } = "דיסק פורק";
        public string AnsNewDisk { get; set; } = "דיסק הותקן";
        public string AnsCD { get; set; } = "אין במחשב CD";
        public string AnsScreenSize { get; set; } = "14";
        public string AnsScrews { get; set; } = "כל הברגים מוברגים עד הסוף";
        public string AnsClean { get; set; } = "המחשב נקי";
        public string AnsExtractAppearance { get; set; } = "מראה חיצוני תקין";
        public string AnsTouchScreen { get; set; } = "לא";
        public string Notes1 { get; set; } = "";
        public string Notes2 { get; set; } = "";

        private ComboBox cmbPlastic, cmbOldDisk, cmbNewDisk, cmbCD, cmbScreenSize, cmbScrews, cmbClean, cmbAppearance, cmbTouch;
        private TextBox txtNotes1, txtNotes2;

        public QaQuestionnaireForm()
        {
            this.Text = "שאלון QA פיזי";
            this.Size = new Size(600, 750);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.None;
            this.BackColor = Color.FromArgb(28, 28, 28);
            this.ForeColor = Color.White;
            this.RightToLeft = RightToLeft.No;

            // Border
            this.Paint += (s, e) => {
                using (Pen p = new Pen(Color.FromArgb(70, 70, 70), 2))
                    e.Graphics.DrawRectangle(p, 0, 0, this.Width - 1, this.Height - 1);
            };

            Panel titleBar = new Panel() { Dock = DockStyle.Top, Height = 40, BackColor = Color.FromArgb(20, 20, 20), Cursor = Cursors.Default };
            titleBar.MouseDown += (s, e) => {
                if (e.Button == MouseButtons.Left) {
                    ReleaseCapture();
                    SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
                }
            };
            Label lblTitle = new Label() { Text = "שאלון בדיקות חומרה פיזיות - QA", ForeColor = Color.Cyan, Font = new Font("Segoe UI", 12, FontStyle.Bold), Location = new Point(10, 8), AutoSize = true, BackColor = Color.Transparent };
            titleBar.Controls.Add(lblTitle);
            
            Label lblClose = new Label() { Text = "✕", Font = new Font("Segoe UI", 12, FontStyle.Bold), Location = new Point(this.Width - 30, 8), AutoSize = true, ForeColor = Color.DarkGray, Cursor = Cursors.Hand, BackColor = Color.Transparent };
            lblClose.MouseEnter += (s, e) => lblClose.ForeColor = Color.LightCoral;
            lblClose.MouseLeave += (s, e) => lblClose.ForeColor = Color.DarkGray;
            lblClose.Click += (s, e) => { this.DialogResult = DialogResult.Cancel; this.Close(); };
            titleBar.Controls.Add(lblClose);
            this.Controls.Add(titleBar);

            Panel contentPanel = new Panel() { Dock = DockStyle.Fill, AutoScroll = true };
            contentPanel.Padding = new Padding(0, 40, 0, 0); // below titlebar

            Font labelFont = new Font("Segoe UI", 10, FontStyle.Bold);
            Font inputFont = new Font("Segoe UI", 10);
            int yPos = 20;

            void AddRow(string labelText, Control inputControl)
            {
                Label lbl = new Label() { Text = labelText, Location = new Point(280, yPos + 3), AutoSize = true, Font = labelFont, RightToLeft = RightToLeft.Yes };
                inputControl.Location = new Point(40, yPos);
                inputControl.Width = 220;
                inputControl.Font = inputFont;
                if (inputControl is ComboBox cb)
                {
                    cb.DropDownStyle = ComboBoxStyle.DropDownList;
                    cb.BackColor = Color.FromArgb(40, 40, 42);
                    cb.ForeColor = Color.White;
                    cb.FlatStyle = FlatStyle.Flat;
                    cb.RightToLeft = RightToLeft.Yes;
                }
                if (inputControl is TextBox tb)
                {
                    tb.BackColor = Color.FromArgb(40, 40, 42);
                    tb.ForeColor = Color.White;
                    tb.BorderStyle = BorderStyle.FixedSingle;
                    tb.RightToLeft = RightToLeft.Yes;
                }
                
                contentPanel.Controls.Add(lbl);
                contentPanel.Controls.Add(inputControl);
                yPos += 45;
            }

            cmbPlastic = new ComboBox(); cmbPlastic.Items.AddRange(new[]{"כן, המחשב שלם", "לא, יש שברים"}); cmbPlastic.SelectedIndex = 0;
            cmbOldDisk = new ComboBox(); cmbOldDisk.Items.AddRange(new[]{"דיסק פורק", "לא רלוונטי"}); cmbOldDisk.SelectedIndex = 0;
            cmbNewDisk = new ComboBox(); cmbNewDisk.Items.AddRange(new[]{"דיסק הותקן", "לא רלוונטי"}); cmbNewDisk.SelectedIndex = 0;
            cmbCD = new ComboBox(); cmbCD.Items.AddRange(new[]{"אין במחשב CD", "יש CD ונפתח תקין", "יש CD לא תקין"}); cmbCD.SelectedIndex = 0;
            cmbScreenSize = new ComboBox(); cmbScreenSize.Items.AddRange(new[]{"14", "15.6", "13.3", "12.5", "17"}); cmbScreenSize.SelectedIndex = 0;
            cmbTouch = new ComboBox(); cmbTouch.Items.AddRange(new[]{"לא", "כן"}); cmbTouch.SelectedIndex = 0;
            cmbScrews = new ComboBox(); cmbScrews.Items.AddRange(new[]{"כל הברגים מוברגים עד הסוף", "חסרים ברגים"}); cmbScrews.SelectedIndex = 0;
            cmbClean = new ComboBox(); cmbClean.Items.AddRange(new[]{"המחשב נקי", "המחשב מלוכלך"}); cmbClean.SelectedIndex = 0;
            cmbAppearance = new ComboBox(); cmbAppearance.Items.AddRange(new[]{"מראה חיצוני תקין", "פגמים במראה ניכרים"}); cmbAppearance.SelectedIndex = 0;
            
            txtNotes1 = new TextBox();
            txtNotes2 = new TextBox();

            AddRow("2. האם המחשב שלם, ללא שברים בפלסטיקה?", cmbPlastic);
            AddRow("3. פירוק דיסק ישן:", cmbOldDisk);
            AddRow("4. התקנת דיסק מעודכן מ\"מחשבים\":", cmbNewDisk);
            AddRow("11. במידה וקיים מתקן דיסק (CD):", cmbCD);
            AddRow("12. גודל המסך (אינטש):", cmbScreenSize);
            AddRow("12.1 מסך מגע:", cmbTouch);
            AddRow("28. ברגים:", cmbScrews);
            AddRow("29. ניקיון המחשב:", cmbClean);
            AddRow("30. מראה חיצוני:", cmbAppearance);
            AddRow("31. הערות לגבי המחשב (למשל מקלדת באנגלית):", txtNotes1);
            AddRow("32. מידע כללי נוסף:", txtNotes2);

            Button btnReview = new Button() {
                Text = "שמור ועבור לסקירה",
                Location = new Point(40, yPos + 10),
                Size = new Size(220, 45),
                BackColor = Color.FromArgb(32, 160, 100),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnReview.FlatAppearance.BorderSize = 0;
            btnReview.Click += BtnReview_Click;
            contentPanel.Controls.Add(btnReview);

            this.Controls.Add(contentPanel);
            contentPanel.BringToFront();
        }

        private void BtnReview_Click(object sender, EventArgs e)
        {
            AnsPlastic = cmbPlastic.SelectedItem.ToString();
            AnsOldDisk = cmbOldDisk.SelectedItem.ToString();
            AnsNewDisk = cmbNewDisk.SelectedItem.ToString();
            AnsCD = cmbCD.SelectedItem.ToString();
            AnsScreenSize = cmbScreenSize.SelectedItem.ToString();
            AnsScrews = cmbScrews.SelectedItem.ToString();
            AnsClean = cmbClean.SelectedItem.ToString();
            AnsExtractAppearance = cmbAppearance.SelectedItem.ToString();
            AnsTouchScreen = cmbTouch.SelectedItem.ToString();
            Notes1 = txtNotes1.Text.Trim();
            Notes2 = txtNotes2.Text.Trim();

            DialogResult res = MessageBox.Show(
                "האם אתה בטוח שאלו התשובות הנכונות?\nלחץ 'כן' כדי לאשר ולהפיק דוח, או 'לא' כדי לחזור ולתקן את המידע כאן.",
                "אישור נתונים", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (res == DialogResult.Yes)
            {
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            // else stay on form to edit
        }
    }
}
