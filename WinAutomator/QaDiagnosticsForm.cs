using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace WinAutomator
{
    public class QaDiagnosticsForm : Form
    {
        public DialogResult MicResult { get; private set; } = DialogResult.None;
        public DialogResult SpeakerResult { get; private set; } = DialogResult.None;
        public DialogResult CameraResult { get; private set; } = DialogResult.None;
        public DialogResult KeyboardResult { get; private set; } = DialogResult.None;
        public DialogResult TrackpadResult { get; private set; } = DialogResult.None;
        public DialogResult UsbResult { get; private set; } = DialogResult.None;

        private ListView listView;
        private Button btnStartAll;
        private Button btnRetest;
        private Button btnFinish;
        private ProgressBar progressQA;
        private Label lblActiveTest;

        private class TestItem
        {
            public string Id;
            public string DisplayName;
            public DialogResult Status = DialogResult.None;
            public Func<Task<DialogResult>> TestAction;
        }

        private List<TestItem> tests;

        public QaDiagnosticsForm()
        {
            this.Text = "QA Diagnostics - בדיקות מעבדה";
            this.Size = new Size(500, 550);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.FromArgb(30, 30, 30);
            this.ForeColor = Color.White;
            this.RightToLeft = RightToLeft.Yes;
            this.RightToLeftLayout = true;

            Label title = new Label() {
                Text = "לוח בקרה - בדיקות חומרה (QA)",
                Dock = DockStyle.Top,
                Height = 60,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 16, FontStyle.Bold),
                ForeColor = Color.Cyan
            };
            this.Controls.Add(title);

            listView = new ListView() {
                Location = new Point(30, 70),
                Size = new Size(420, 250),
                View = View.Details,
                FullRowSelect = true,
                BackColor = Color.FromArgb(40, 40, 40),
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 11),
                MultiSelect = false
            };
            listView.Columns.Add("רכיב", 250);
            listView.Columns.Add("סטטוס", 150);
            listView.SelectedIndexChanged += ListView_SelectedIndexChanged;
            this.Controls.Add(listView);

            lblActiveTest = new Label() {
                Text = "ממתין להתחלת מערך בדיקות...",
                Location = new Point(30, 330),
                Size = new Size(420, 25),
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 10, FontStyle.Italic)
            };
            this.Controls.Add(lblActiveTest);

            progressQA = new ProgressBar() {
                Location = new Point(30, 360),
                Size = new Size(420, 20),
                Style = ProgressBarStyle.Blocks,
                Maximum = 5
            };
            this.Controls.Add(progressQA);

            btnStartAll = new Button() {
                Text = "▶ התחל אבחון רציף",
                Location = new Point(30, 400),
                Size = new Size(130, 45),
                BackColor = Color.SteelBlue,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btnStartAll.FlatAppearance.BorderSize = 0;
            btnStartAll.Click += async (s, e) => await RunAllTests();
            this.Controls.Add(btnStartAll);

            btnRetest = new Button() {
                Text = "↻ בדיקה חוזרת",
                Location = new Point(175, 400),
                Size = new Size(130, 45),
                BackColor = Color.Goldenrod,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Enabled = false // Only enabled when selecting an item after run
            };
            btnRetest.FlatAppearance.BorderSize = 0;
            btnRetest.Click += async (s, e) => await RunSelectedTest();
            this.Controls.Add(btnRetest);

            btnFinish = new Button() {
                Text = "✓ סיים וייצא דו\"ח",
                Location = new Point(320, 400),
                Size = new Size(130, 45),
                BackColor = Color.SeaGreen,
                FlatStyle = FlatStyle.Flat,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Enabled = false
            };
            btnFinish.FlatAppearance.BorderSize = 0;
            btnFinish.Click += (s, e) => FinalizeResults();
            this.Controls.Add(btnFinish);

            SetupTests();
            RefreshList();
        }

        private void SetupTests()
        {
            tests = new List<TestItem>
            {
                new TestItem { Id = "Mic", DisplayName = "מיקרופון", TestAction = RunMicTest },
                new TestItem { Id = "Speaker", DisplayName = "רמקולים (סטריאו)", TestAction = RunSpeakerTest },
                new TestItem { Id = "Cam", DisplayName = "מצלמה", TestAction = RunCamTest },
                new TestItem { Id = "KB", DisplayName = "מקלדת", TestAction = RunKbTest },
                new TestItem { Id = "TP", DisplayName = "משטח מגע / עכבר", TestAction = RunTrackpadTest },
                new TestItem { Id = "USB", DisplayName = "שקעי USB", TestAction = RunUsbTest }
            };
        }

        private void RefreshList()
        {
            if (InvokeRequired) { Invoke(new Action(RefreshList)); return; }
            listView.Items.Clear();
            foreach (var t in tests)
            {
                ListViewItem item = new ListViewItem(t.DisplayName);
                string statusTxt = "ממתין ⏳";
                if (t.Status == DialogResult.Yes) { statusTxt = "תקין ✓"; item.ForeColor = Color.LimeGreen; }
                else if (t.Status == DialogResult.No) { statusTxt = "נכשל ✗"; item.ForeColor = Color.LightCoral; }
                else if (t.Status == DialogResult.Ignore) { statusTxt = "דולג ⏭"; item.ForeColor = Color.Goldenrod; }

                item.SubItems.Add(statusTxt);
                listView.Items.Add(item);
            }
        }

        private void ListView_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (btnFinish.Enabled && listView.SelectedIndices.Count > 0)
            {
                btnRetest.Enabled = true;
            }
            else
            {
                btnRetest.Enabled = false;
            }
        }

        private void LockForTesting(bool isLock)
        {
            if (InvokeRequired) { Invoke(new Action(() => LockForTesting(isLock))); return; }
            btnStartAll.Enabled = !isLock;
            btnRetest.Enabled = !isLock && listView.SelectedIndices.Count > 0;
            listView.Enabled = !isLock;
        }

        private async Task RunAllTests()
        {
            LockForTesting(true);
            btnFinish.Enabled = false;
            progressQA.Value = 0;

            for (int i = 0; i < tests.Count; i++)
            {
                var t = tests[i];
                lblActiveTest.Text = $"בודק כעת: {t.DisplayName}...";
                t.Status = await t.TestAction();
                RefreshList();
                progressQA.Value = i + 1;
            }

            lblActiveTest.Text = "אבחון רציף הסתיים. ניתן לבצע בדיקה חוזרת או לסיים.";
            btnFinish.Enabled = true;
            LockForTesting(false);
        }

        private async Task RunSelectedTest()
        {
            if (listView.SelectedIndices.Count == 0) return;
            int idx = listView.SelectedIndices[0];
            var t = tests[idx];

            LockForTesting(true);
            lblActiveTest.Text = $"בודק שוב (בדיקה חוזרת): {t.DisplayName}...";
            
            DialogResult newRes = await t.TestAction();
            
            // If it failed again, prompt warning
            if (newRes == DialogResult.No)
            {
                MessageBox.Show($"בדיקה חוזרת נכשלה שוב עבור רכיב: {t.DisplayName}!\nהסטטוס נותר 'נכשל'.", "אזהרת חומרה", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            
            t.Status = newRes;

            RefreshList();
            lblActiveTest.Text = "ניתן לבצע בדיקה חוזרת על רכיבים אחרים או לסיים.";
            LockForTesting(false);
            btnFinish.Enabled = true; // explicitly enable to ensure workflow continues
        }

        private void FinalizeResults()
        {
            MicResult = tests[0].Status;
            SpeakerResult = tests[1].Status;
            CameraResult = tests[2].Status;
            KeyboardResult = tests[3].Status;
            TrackpadResult = tests[4].Status;
            UsbResult = tests[5].Status;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        // --- Hardware Logic Wrappers ---

        private async Task<DialogResult> RunMicTest()
        {
            MessageBox.Show("נבדוק כעת מיקרופון. כשתלחץ אישור תחל הקלטה רצופה של 20 שניות. דבר למחשב.", "בדיקות סאונד", MessageBoxButtons.OK, MessageBoxIcon.Information);
            AutomationLogic.StartMicRecording();
            await Task.Delay(20000);
            AutomationLogic.StopAndPlayMic();
            return MessageBox.Show("האם המיקרופון תקין והקלטת את עצמך קורא וברור?", "מבדק סאונד", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        }

        private async Task<DialogResult> RunSpeakerTest()
        {
            MessageBox.Show("נבדוק כעת את הרמקולים. אתה תשמע צליל 'ביפ' קודם משמאל ואז מימין.", "בדיקת רמקולים", MessageBoxButtons.OK, MessageBoxIcon.Information);
            await Task.Run(() => AutomationLogic.PlayStereoTest(true));
            await Task.Run(() => AutomationLogic.PlayStereoTest(false));
            return MessageBox.Show("האם שמעת צליל בשני הרמקולים (שמאל וימין) בנפרד?", "מבדק סטריאו", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        }

        private async Task<DialogResult> RunCamTest()
        {
            MessageBox.Show("נבדוק כעת את המצלמה. המצלמה תיפתח ל-30 שניות ותיסגר.", "בדיקות מצלמה", MessageBoxButtons.OK, MessageBoxIcon.Information);
            AutomationLogic.OpenCameraFor30Secs();
            await Task.Delay(30000);
            AutomationLogic.CloseCamera();
            return MessageBox.Show("האם המצלמה תקינה וראית תמונה?", "מבדק מצלמה", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        }

        private async Task<DialogResult> RunKbTest()
        {
            using (var kbForm = new QaKeyboardForm())
            {
                // To keep the user focused, if main form dimmed, this will sit nicely on top.
                kbForm.ShowDialog();
            }
            return MessageBox.Show("מבדק: הבדיקה הושלמה חזותית - האם המקלדת טקטלית ותקינה?", "מבדק מקלדת", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        }

        private async Task<DialogResult> RunTrackpadTest()
        {
            return await Task.Run(() => MessageBox.Show("אנא ודא כעת שימוש במשטח מגע העכבר (Trackpad). תקין?", "מבדק עכבר", MessageBoxButtons.YesNo, MessageBoxIcon.Question));
        }

        private async Task<DialogResult> RunUsbTest()
        {
            DialogResult res = DialogResult.None;
            using (var usbForm = new QaUsbForm())
            {
                res = usbForm.ShowDialog(this);
            }
            if (res == DialogResult.Ignore) 
            {
                MessageBox.Show("בדיקת USB דולגה לבקשתך.", "בדיקת USB", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            return res;
        }
    }
}
