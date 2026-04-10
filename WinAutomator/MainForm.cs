using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace WinAutomator
{
    public class MainForm : Form
    {
        private int currentPhase;
        private string techName;
        private string serialNum;
        private string cpuGen;
        private bool skipHostname;
        private bool isFullAuto;

        // UI Controls
        private TextBox txtTech;
        private TextBox txtSerial;
        private TextBox txtCpu;
        private CheckBox chkSkipHostname;
        private Button btnFullAuto;
        private Button btnSemiAuto;
        private ProgressBar progressBar;
        private Label lblStatus;
        private Panel headerPanel;

        // === NEW: Timer System ===
        private Stopwatch processStopwatch;
        private System.Windows.Forms.Timer elapsedTimer;
        private Label lblTimer;

        // === NEW: Detailed Log System ===
        private RichTextBox rtbLog;
        private Panel logPanel;
        private List<string> logEntries = new List<string>();

        public MainForm(int phase, string tech, string serial, string cpu, bool isFullAutoCfg, bool skipHostnameCfg = false)
        {
            this.currentPhase = phase;
            this.techName = tech;
            this.serialNum = serial;
            this.cpuGen = cpu;
            this.isFullAuto = isFullAutoCfg;
            this.skipHostname = skipHostnameCfg;

            InitializeDarkUI();
            InitializeTimerUI();
            InitializeLogPanel();

            if (currentPhase == 1 && string.IsNullOrEmpty(this.techName))
            {
                // First time fresh start!
                ShowInputFields();
            }
            else
            {
                // Resuming from Restart! (Or looping phase 1)
                LockUIForAutomation();
                StartGlobalTimer();
                
                if (currentPhase == 1) ExecutePhase1();
                else if (currentPhase == 2) ExecutePhase2();
                else if (currentPhase == 3) ExecutePhase3();
            }
        }

        private void InitializeDarkUI()
        {
            this.Text = "WinAutomator - 8.0";
            this.Size = new Size(550, 680);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.BackColor = Color.FromArgb(30, 30, 30);
            this.ForeColor = Color.White;
            this.RightToLeft = RightToLeft.Yes;
            this.RightToLeftLayout = true;

            // Custom Glowing Header
            headerPanel = new Panel() { Dock = DockStyle.Top, Height = 90, BackColor = Color.FromArgb(20, 20, 20) };
            headerPanel.Paint += HeaderPanel_Paint;
            this.Controls.Add(headerPanel);

            // Status Label (Real-time feedback)
            lblStatus = new Label() { 
                Text = "ממתין לפקודה...", 
                Location = new Point(10, 270), 
                Width = 510, 
                Height = 25,
                TextAlign = ContentAlignment.TopCenter, 
                AutoSize = false, 
                Font = new Font("Segoe UI", 11, FontStyle.Bold),
                ForeColor = Color.Cyan
            };

            // Progress Bar (Prominent Cyan)
            progressBar = new ProgressBar() { 
                Location = new Point(40, 305),
                Width = 450, 
                Height = 35, // Larger bar
                Visible = false 
            };

            this.Controls.Add(progressBar);
            this.Controls.Add(lblStatus);
        }

        // === NEW: Timer UI Initialization ===
        private void InitializeTimerUI()
        {
            processStopwatch = new Stopwatch();
            
            lblTimer = new Label()
            {
                Text = "⏱ 00:00:00",
                Location = new Point(40, 345), // Moved to align with toggle button
                Width = 200,
                Height = 30,
                Font = new Font("Consolas", 14, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 200, 255),
                TextAlign = ContentAlignment.MiddleLeft,
                Visible = false,
                RightToLeft = RightToLeft.No
            };
            this.Controls.Add(lblTimer);

            elapsedTimer = new System.Windows.Forms.Timer() { Interval = 1000 };
            elapsedTimer.Tick += (s, e) =>
            {
                if (processStopwatch.IsRunning)
                {
                    TimeSpan ts = processStopwatch.Elapsed;
                    lblTimer.Text = $"⏱ {ts.Hours:D2}:{ts.Minutes:D2}:{ts.Seconds:D2}";
                }
            };
        }

        private void StartGlobalTimer()
        {
            lblTimer.Visible = true;
            if (!processStopwatch.IsRunning) processStopwatch.Start(); // Start only if not already running (for resume)
            elapsedTimer.Start();
        }

        private void StopGlobalTimer()
        {
            processStopwatch.Stop();
            elapsedTimer.Stop();
        }

        private string GetElapsedTimeString()
        {
            TimeSpan ts = processStopwatch.Elapsed;
            return $"{ts.Hours:D2}:{ts.Minutes:D2}:{ts.Seconds:D2}";
        }

        // === NEW: Detailed Log Panel ===
        private void InitializeLogPanel()
        {
            // Log panel
            logPanel = new Panel()
            {
                Location = new Point(20, 385), // Higher up, under the timer
                Size = new Size(495, 200),
                BackColor = Color.FromArgb(18, 18, 18),
                BorderStyle = BorderStyle.FixedSingle,
                Visible = false
            };

            rtbLog = new RichTextBox()
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(18, 18, 18),
                ForeColor = Color.FromArgb(180, 220, 180),
                Font = new Font("Consolas", 9),
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                RightToLeft = RightToLeft.No,
                WordWrap = true
            };
            logPanel.Controls.Add(rtbLog);

            this.Controls.Add(logPanel);
        }

        private void AppendLog(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            string entry = $"[{timestamp}] {message}";
            logEntries.Add(entry);

            if (InvokeRequired) 
            { 
                Invoke(new Action(() => AppendLog(message))); 
                return; 
            }

            rtbLog.AppendText(entry + "\n");
            rtbLog.ScrollToCaret();
        }

        private void HeaderPanel_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
            
            string title = "תוכנה להתקנת מחשב";
            string prefix = "נבנה על ידי ";
            string nameWord1 = "יקיר";
            string nameWord2 = "לביא";
            
            Font fontTitle = new Font("Segoe UI", 13, FontStyle.Regular);
            Font fontCredit = new Font("Segoe UI", 20, FontStyle.Bold);

            // --- Line 1: Title (centered, static gray) ---
            SizeF titleSize = e.Graphics.MeasureString(title, fontTitle);
            float titleX = (headerPanel.Width - titleSize.Width) / 2f;
            e.Graphics.DrawString(title, fontTitle, Brushes.Gray, new PointF(titleX, 6));

            // --- Line 2: Credit (draw as 3 segments: prefix + word1 + space + word2) ---
            // Measure each segment to position them correctly
            // In RTL: visual order is  לביא  יקיר  נבנה על ידי  (right to left)
            // But DrawString with RTL handles this for us if we draw segments in logical order
            
            // Use StringFormat for proper RTL rendering
            StringFormat sfRtl = new StringFormat(StringFormat.GenericTypographic);
            sfRtl.FormatFlags |= StringFormatFlags.DirectionRightToLeft;
            
            // Measure the full credit line for centering
            string fullCredit = prefix + nameWord1 + " " + nameWord2;
            SizeF fullSize = e.Graphics.MeasureString(fullCredit, fontCredit);
            float creditX = (headerPanel.Width - fullSize.Width) / 2f;
            float creditY = 40;
            
            // Draw the prefix "נבנה על ידי " in static gray
            e.Graphics.DrawString(prefix, fontCredit, Brushes.DarkGray, new PointF(creditX, creditY));
            
            // Measure prefix width to know where the name starts
            SizeF prefixSize = e.Graphics.MeasureString(prefix, fontCredit);
            float nameStartX = creditX + prefixSize.Width;
            
            // Draw name word 1: "יקיר" with static glow
            Color color1 = Color.Cyan;
            
            // Glow layer for word 1
            using (SolidBrush glowBrush = new SolidBrush(Color.FromArgb(40, color1)))
            {
                for (int r = 1; r <= 2; r++)
                {
                    e.Graphics.DrawString(nameWord1, fontCredit, glowBrush, nameStartX - r, creditY - r);
                    e.Graphics.DrawString(nameWord1, fontCredit, glowBrush, nameStartX + r, creditY + r);
                }
            }
            using (SolidBrush textBrush = new SolidBrush(color1))
            {
                e.Graphics.DrawString(nameWord1, fontCredit, textBrush, nameStartX, creditY);
            }
            
            // Measure word1 to position word2
            SizeF word1Size = e.Graphics.MeasureString(nameWord1 + " ", fontCredit);
            float word2X = nameStartX + word1Size.Width;
            
            // Draw name word 2: "לביא" with static glow
            Color color2 = Color.DodgerBlue;
            
            // Glow layer for word 2
            using (SolidBrush glowBrush = new SolidBrush(Color.FromArgb(40, color2)))
            {
                for (int r = 1; r <= 2; r++)
                {
                    e.Graphics.DrawString(nameWord2, fontCredit, glowBrush, word2X - r, creditY - r);
                    e.Graphics.DrawString(nameWord2, fontCredit, glowBrush, word2X + r, creditY + r);
                }
            }
            using (SolidBrush textBrush = new SolidBrush(color2))
            {
                e.Graphics.DrawString(nameWord2, fontCredit, textBrush, word2X, creditY);
            }
        }



        private void ShowInputFields()
        {
            Font f = new Font("Segoe UI", 11);
            
            Label l1 = new Label() { Text = "שם הטכנאי:", Location = new Point(360, 120), AutoSize = true, Font = f };
            txtTech = new TextBox() { Location = new Point(90, 117), Width = 250, Font = f, BackColor = Color.FromArgb(50,50,50), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };

            Label l2 = new Label() { Text = "מספר סידורי:", Location = new Point(360, 170), AutoSize = true, Font = f };
            txtSerial = new TextBox() { Location = new Point(90, 167), Width = 250, Font = f, BackColor = Color.FromArgb(50,50,50), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };

            Label l3 = new Label() { Text = "דור מעבד:", Location = new Point(360, 220), AutoSize = true, Font = f };
            txtCpu = new TextBox() { Location = new Point(90, 217), Width = 250, Font = f, BackColor = Color.FromArgb(50,50,50), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };

            chkSkipHostname = new CheckBox() { 
                Text = "דלג על שינוי שם המחשב (מצב בדיקה)", 
                Location = new Point(90, 255), 
                Width = 350, 
                Font = new Font("Segoe UI", 10, FontStyle.Bold), 
                ForeColor = Color.FromArgb(255, 160, 0) 
            };

            btnFullAuto = new Button() { Text = "אוטומציה מלאה", Location = new Point(270, 300), Width = 150, Height = 40, FlatStyle = FlatStyle.Flat, BackColor = Color.SeaGreen, ForeColor = Color.White, Font = f, Cursor = Cursors.Hand };
            btnFullAuto.FlatAppearance.BorderSize = 0;
            btnFullAuto.Click += (s, e) => StartFresh(true);

            btnSemiAuto = new Button() { Text = "חצי אוטומטי", Location = new Point(90, 300), Width = 150, Height = 40, FlatStyle = FlatStyle.Flat, BackColor = Color.SteelBlue, ForeColor = Color.White, Font = f, Cursor = Cursors.Hand };
            btnSemiAuto.FlatAppearance.BorderSize = 0;
            btnSemiAuto.Click += (s, e) => StartFresh(false);

            this.Controls.Add(l1); this.Controls.Add(txtTech);
            this.Controls.Add(l2); this.Controls.Add(txtSerial);
            this.Controls.Add(l3); this.Controls.Add(txtCpu);
            this.Controls.Add(chkSkipHostname);
            this.Controls.Add(btnFullAuto); this.Controls.Add(btnSemiAuto);
        }

        private void StartFresh(bool fullAuto)
        {
            if(string.IsNullOrWhiteSpace(txtTech.Text) || string.IsNullOrWhiteSpace(txtSerial.Text) || string.IsNullOrWhiteSpace(txtCpu.Text))
            {
                MessageBox.Show("נא למלא את כל 3 השדות טרם תחילת האוטומציה.", "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            this.techName = txtTech.Text.Trim();
            this.serialNum = txtSerial.Text.Trim();
            this.cpuGen = txtCpu.Text.Trim();
            this.isFullAuto = fullAuto;
            this.skipHostname = chkSkipHostname.Checked;

            // Hide the Inputs to show the loading screen
            foreach (Control c in this.Controls) {
                if (c is TextBox || c is Button || c is CheckBox || (c is Label && c.Parent != headerPanel)) c.Visible = false;
            }

            LockUIForAutomation();
            StartGlobalTimer();
            AppendLog($"התחלת תהליך ({(fullAuto ? "אוטומציה מלאה" : "חצי אוטומטי")})");
            AppendLog($"טכנאי: {techName} | סיריאלי: {serialNum} | דור CPU: {cpuGen}");
            ExecutePhase1();
        }

        private void LockUIForAutomation()
        {
            progressBar.Visible = true;
            // Indeterminate Marquee bar
            progressBar.Style = ProgressBarStyle.Marquee;
            progressBar.MarqueeAnimationSpeed = 30;

            // Show log panel permanently
            logPanel.Visible = true;
        }

        private void UpdateStatus(string message)
        {
            if (InvokeRequired) { Invoke(new Action(() => UpdateStatus(message))); return; }
            lblStatus.Text = message;
            AppendLog(message);
        }

        // ======================= PHASE 1 =======================
        private async void ExecutePhase1()
        {
            AppendLog("═══ פאזה 1: הכנת דיסק ותיקון ליבה ═══");

            UpdateStatus("פאזה 1: סורק ומרחיב דיסק (במידה וניתן)...");
            await Task.Run(() => AutomationLogic.ExtendCDrive());
            AppendLog("הרחבת דיסק C הושלמה");

            UpdateStatus("פאזה 1: מריץ פקודת DISM לשיקום מערכת. נא להמתין...");
            bool dismSuccess = await Task.Run(() => AutomationLogic.RunDismCommand(this.isFullAuto));
            AppendLog($"DISM הסתיים (הצלחה: {dismSuccess})");

            if (!this.isFullAuto)
            {
                DialogResult res = MessageBox.Show("האם תהליך ה-DISM הסתיים בהצלחה בלי שגיאות אדומות?", "בקרת DISM חציונית", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (res == DialogResult.No)
                {
                    AppendLog("⚠ כשל DISM דווח ע\"י הטכנאי – מתחיל לופ ריסטארט");
                    UpdateStatus("כשל DISM דווח! מקנפג חזרה ללופ פאזה 1 ומבצע הפעלה מחדש...");
                    await Task.Delay(2000);
                    AutomationLogic.SetupAutoResumeAndRestart(1, techName, serialNum, cpuGen, isFullAuto, skipHostname);
                    return; // Stops here, restarts PC.
                }
            }
            
            // Success! Move to Phase 2
            AppendLog("✓ פאזה 1 הושלמה בהצלחה – מכין ריסטארט לפאזה 2");
            UpdateStatus("פאזה 1 הסתיימה! מכין מחשב להפעלה מחדש (לפני פאזה 2)...");
            await Task.Delay(2000);
            AutomationLogic.SetupAutoResumeAndRestart(2, techName, serialNum, cpuGen, isFullAuto, skipHostname);
        }

        // ======================= PHASE 2 =======================
        private async void ExecutePhase2()
        {
            AppendLog("═══ פאזה 2: רשיונות, דרייברים ועדכונים ═══");

            if (skipHostname)
            {
                UpdateStatus("פאזה 2: דילוג על שינוי שם המחשב לפי בקשת המשתמש...");
                AppendLog("דילוג על Hostname (בקשת משתמש)");
            }
            else
            {
                UpdateStatus("פאזה 2: מעדכן שם מחשב (Hostname)...");
                await Task.Run(() => AutomationLogic.ChangeHostname(serialNum));
                AppendLog($"Hostname שונה ל: {serialNum}");
            }

            UpdateStatus("פאזה 2: מבצע אקטיבציה ל-Windows 11...");
            await Task.Run(() => AutomationLogic.ActivateWindows());
            AppendLog("אקטיבציית Windows הושלמה");

            UpdateStatus("פאזה 2: מבצע אקטיבציה ל-Office 2021...");
            await Task.Run(() => AutomationLogic.ActivateOffice());
            AppendLog("אקטיבציית Office הושלמה");

            // Heavy Updates
            await Task.Run(() => AutomationLogic.RunOemUpdates(UpdateStatus));
            AppendLog("עדכוני OEM הושלמו");
            await Task.Run(() => AutomationLogic.RunWindowsUpdates(UpdateStatus));
            AppendLog("בקשת עדכוני Windows נשלחה. ממתין...");

            UpdateStatus("ממתין 60 שניות כדי לאפשר לעדכוני Windows לרדת ברקע...");
            await Task.Delay(60000);
            AppendLog("עדכוני Windows סיימו המתנה של דקה לחילוץ חבילות");

            // Success! Move to Phase 3
            AppendLog("✓ פאזה 2 הושלמה בהצלחה – מכין ריסטארט לפאזה 3");
            UpdateStatus("פאזה 2 הסתיימה בהצלחה! השלמנו מאסת עדכונים, מאתחל (לפאזה 3)...");
            await Task.Delay(2000);
            AutomationLogic.SetupAutoResumeAndRestart(3, techName, serialNum, cpuGen, isFullAuto, skipHostname);
        }

        // ======================= PHASE 3 (QA) =======================
        private async void ExecutePhase3()
        {
            AppendLog("═══ פאזה 3: אבחון חומרה ובדיקות QA ═══");
            progressBar.Style = ProgressBarStyle.Blocks;
            progressBar.Value = 5;
            
            // --- Defrag / Optimize SSD ---
            UpdateStatus("מבצע אופטימיזציה לכונן (TRIM/Defrag)...");
            string defragOutput = await Task.Run(() => AutomationLogic.OptimizeDrive(UpdateStatus));
            AppendLog("תוצאת Defrag/Optimize:\n" + defragOutput.Trim());
            progressBar.Value = 10;

            UpdateStatus("מפיק דו\"ח בריאות סוללה ומשדר ל-API ב-Base44...");
            await AutomationLogic.PerformBatteryReportAndApi(techName, serialNum, cpuGen, UpdateStatus);
            AppendLog("דוח סוללה שוגר ל-API");

            UpdateStatus("סיימנו התקנות! מכין סביבת בדיקות (QA Diagnostics)...");
            AppendLog("♫ משמיע צליל סיום התקנות...");
            await Task.Run(() => AutomationLogic.PlayVictorySound());
            
            string summaryMsg = 
                "התהליך האוטומטי מאחורי הקלעים הושלם!\n\n" +
                "הפעולות שבוצעו במחשב לאורך ההפעלה:\n" +
                "============================\n" +
                "✓ שחזור רכיבי מערכת ופגמים (DISM)\n" +
                "✓ הרחבת מחיצת אחסון של כונן C\n" +
                "✓ סריקה והתקנת דרייברים מהיצרן\n" +
                "✓ משיכת עדכוני אבטחה (Windows Update)\n" +
                "✓ ביצוע אקטיבציה של הווינדוס ואופיס\n" +
                "✓ ביצוע אופטימיזציה לסוג הכונן (TRIM/Defrag)\n" +
                "✓ ניקוי מקיף של קבצי זבל, שאריות ומטמון\n\n" +
                "כעת אנו עוברים לשלב הסופי: בדיקות חומרה (QA Diagnosics).";
                
            MessageBox.Show(summaryMsg, "סיכום משימות אוטומציה", MessageBoxButtons.OK, MessageBoxIcon.Information);

            UpdateStatus("פותח את מסך האבחון / QA. נא להמשיך בחלון החדש...");

            DialogResult micRes = DialogResult.None;
            DialogResult stereoRes = DialogResult.None;
            DialogResult camRes = DialogResult.None;
            DialogResult kbRes = DialogResult.None;
            DialogResult tpRes = DialogResult.None;
            DialogResult usbRes = DialogResult.None;

            // Run in UI thread specifically for Form execution and locking
            await Task.Run(() => 
            {
                Invoke(new Action(() => 
                {
                    // Create overlay to dim the screen
                    using (Form overlay = new Form())
                    {
                        overlay.StartPosition = FormStartPosition.Manual;
                        overlay.Location = this.Location;
                        overlay.Size = this.Size;
                        overlay.BackColor = Color.Black;
                        overlay.Opacity = 0.6; // Dim the main form nicely
                        overlay.FormBorderStyle = FormBorderStyle.None;
                        overlay.ShowInTaskbar = false;
                        overlay.Show(this);

                        using (var qaForm = new QaDiagnosticsForm())
                        {
                            qaForm.ShowDialog(overlay);
                            
                            // Capture the results
                            micRes = qaForm.MicResult;
                            stereoRes = qaForm.SpeakerResult;
                            camRes = qaForm.CameraResult;
                            kbRes = qaForm.KeyboardResult;
                            tpRes = qaForm.TrackpadResult;
                            usbRes = qaForm.UsbResult;
                        }

                        overlay.Close();
                    }
                }));
            });

            progressBar.Value = 100;

            // Log the captured results directly since the inner windows dealt with user logging
            AppendLog($"מיקרופון: {(micRes == DialogResult.Yes ? "תקין ✓" : "נכשל ✗")}");
            AppendLog($"רמקולים: {(stereoRes == DialogResult.Yes ? "תקין ✓" : "נכשל ✗")}");
            AppendLog($"מצלמה: {(camRes == DialogResult.Yes ? "תקין ✓" : "נכשל ✗")}");
            AppendLog($"מקלדת: {(kbRes == DialogResult.Yes ? "תקין ✓" : "נכשל ✗")}");
            AppendLog($"משטח מגע: {(tpRes == DialogResult.Yes ? "תקין ✓" : "נכשל ✗")}");
            AppendLog($"כניסות USB: {(usbRes == DialogResult.Yes ? "תקין ✓" : (usbRes == DialogResult.Ignore ? "דולג ⏭" : "נכשל ✗"))}");

            // === STOP TIMER ===
            StopGlobalTimer();
            string totalTime = GetElapsedTimeString();
            AppendLog($"═══ התהליך הושלם! זמן כולל: {totalTime} ═══");

            ShowResultsSummary(micRes, camRes, kbRes, tpRes, usbRes, ssdHealth, stereoRes, totalTime);
        }

        private void ShowResultsSummary(DialogResult mic, DialogResult cam, DialogResult kb, DialogResult tp, DialogResult usb, string ssd, DialogResult stereo, string totalTime)
        {
            if (InvokeRequired) { Invoke(new Action(() => ShowResultsSummary(mic, cam, kb, tp, usb, ssd, stereo, totalTime))); return; }
            
            // Clear existing controls except header
            foreach (Control c in this.Controls) if (c != headerPanel) c.Visible = false;

            // Resize form for results (collapse log if expanded)
            this.Height = 530;

            Panel resPanel = new Panel() { 
                Location = new Point(50, 110), Size = new Size(450, 370), 
                BackColor = Color.FromArgb(40, 40, 40), BorderStyle = BorderStyle.FixedSingle 
            };
            
            Label title = new Label() { 
                Text = "סיכום בדיקות שבוצעו", Dock = DockStyle.Top, Height = 35, 
                TextAlign = ContentAlignment.MiddleCenter, Font = new Font("Segoe UI", 12, FontStyle.Bold) 
            };
            resPanel.Controls.Add(title);

            string[] names = { "מיקרופון", "מצלמה", "מקלדת", "משטח מגע", "שקעי USB", "סטריאו רמקולים" };
            DialogResult[] results = { mic, cam, kb, tp, usb, stereo };

            for (int i = 0; i < names.Length; i++)
            {
                Label lblName = new Label() { Text = names[i], Location = new Point(300, 40 + i * 35), AutoSize = true, Font = new Font("Segoe UI", 10) };
                
                string resTxt = "נכשל ✗";
                Color resCol = Color.LightCoral;
                if (results[i] == DialogResult.Yes) { resTxt = "תקין ✓"; resCol = Color.LimeGreen; }
                else if (results[i] == DialogResult.Ignore) { resTxt = "דולג ⏭"; resCol = Color.Goldenrod; }

                Label lblRes = new Label() { 
                    Text = resTxt, 
                    ForeColor = resCol,
                    Location = new Point(50, 40 + i * 35), AutoSize = true, Font = new Font("Segoe UI", 10, FontStyle.Bold) 
                };
                resPanel.Controls.Add(lblName);
                resPanel.Controls.Add(lblRes);
            }

            // SSD Info
            Label lblSsdTitle = new Label() { Text = "בריאות דיסק:", Location = new Point(300, 40 + names.Length * 35), AutoSize = true, Font = new Font("Segoe UI", 10) };
            Label lblSsdVal = new Label() { Text = ssd, Location = new Point(30, 40 + names.Length * 35), Width = 250, ForeColor = Color.Cyan, Font = new Font("Segoe UI", 9, FontStyle.Bold) };
            resPanel.Controls.Add(lblSsdTitle); resPanel.Controls.Add(lblSsdVal);

            // === Total Time Display ===
            int timeY = 40 + (names.Length + 1) * 35;
            Label lblTimeTitle = new Label() { Text = "⏱ זמן כולל:", Location = new Point(300, timeY), AutoSize = true, Font = new Font("Segoe UI", 10) };
            Label lblTimeVal = new Label() 
            { 
                Text = totalTime, 
                Location = new Point(50, timeY), 
                AutoSize = true, 
                ForeColor = Color.FromArgb(0, 200, 255), 
                Font = new Font("Consolas", 12, FontStyle.Bold),
                RightToLeft = RightToLeft.No
            };
            resPanel.Controls.Add(lblTimeTitle); resPanel.Controls.Add(lblTimeVal);

            // === Show Log Button ===
            Button btnShowLog = new Button()
            {
                Text = "📋 הצג לוג מלא",
                Location = new Point(20, 310),
                Width = 190,
                Height = 35,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(55, 55, 55),
                ForeColor = Color.LightGray,
                Font = new Font("Segoe UI", 9),
                Cursor = Cursors.Hand
            };
            btnShowLog.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 80);
            btnShowLog.Click += (s, e) =>
            {
                // Show a new form with the full log
                Form logForm = new Form()
                {
                    Text = "לוג מפורט - WinAutomator",
                    Size = new Size(700, 500),
                    StartPosition = FormStartPosition.CenterParent,
                    BackColor = Color.FromArgb(18, 18, 18),
                    ForeColor = Color.White
                };
                RichTextBox fullLog = new RichTextBox()
                {
                    Dock = DockStyle.Fill,
                    BackColor = Color.FromArgb(18, 18, 18),
                    ForeColor = Color.FromArgb(180, 220, 180),
                    Font = new Font("Consolas", 10),
                    ReadOnly = true,
                    BorderStyle = BorderStyle.None,
                    Text = string.Join("\n", logEntries),
                    RightToLeft = RightToLeft.No
                };
                logForm.Controls.Add(fullLog);
                logForm.ShowDialog();
            };
            resPanel.Controls.Add(btnShowLog);

            Button btnFinish = new Button() { 
                Text = "סיום וסגירה", Location = new Point(230, 310), Width = 190, Height = 35, 
                FlatStyle = FlatStyle.Flat, BackColor = Color.SeaGreen, Cursor = Cursors.Hand 
            };
            btnFinish.FlatAppearance.BorderSize = 0;
            btnFinish.Click += (s, e) => Application.Exit();
            resPanel.Controls.Add(btnFinish);

            this.Controls.Add(resPanel);
        }
    }
}
