using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Text;
using System.Windows.Forms;
using System.Threading.Tasks;
using WinAutomator.Phases;

namespace WinAutomator
{
    public class MainForm : Form
    {
        // === P/Invoke for Draggable Borderless Window ===
        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        // === Context & State ===
        private int currentPhase;
        private readonly AutomationContext context;
        private IAutomationPhase[] allPhases;

        // === Input Controls ===
        private Panel inputPanel;
        private TextBox txtTech;
        private TextBox txtSerial;
        private TextBox txtCpu;
        private CheckBox chkSkipHostname;
        private ComboBox cmbManufacturer;
        private Button btnFullAuto;
        private Button btnSemiAuto;

        // === Automation Controls ===
        private Panel automationPanel;
        private StepperControl stepperControl;
        private Label lblStatus;
        private Label lblManInfo;

        // === Timer System ===
        private Stopwatch processStopwatch;
        private System.Windows.Forms.Timer elapsedTimer;
        private Label lblTimer;

        // === Log System ===
        private RichTextBox rtbLog;
        private Panel logPanel;
        private readonly List<string> logEntries = new();

        // === Shared UI ===
        private Panel headerPanel;

        // =====================================================================
        //  CONSTRUCTOR
        // =====================================================================

        public MainForm(int phase, string tech, string serial, string cpu,
            bool isFullAutoCfg, bool skipHostnameCfg = false, string manufacturer = "")
        {
            this.currentPhase = phase;

            context = new AutomationContext
            {
                TechName = tech,
                SerialNum = serial,
                CpuGen = cpu,
                IsFullAuto = isFullAutoCfg,
                SkipHostname = skipHostnameCfg,
                SelectedManufacturer = manufacturer
            };

            allPhases = new IAutomationPhase[]
            {
                new Phase1_DiskRepair(),
                new Phase2_Updates(),
                new Phase3_QA()
            };

            InitializeForm();
            InitializeHeader();
            InitializeInputScreen();
            InitializeAutomationScreen();

            if (currentPhase == 1 && string.IsNullOrEmpty(context.TechName))
            {
                ShowInputScreen();
            }
            else
            {
                ShowAutomationScreen();
                StartGlobalTimer();
                RunCurrentPhase();
            }
        }

        // =====================================================================
        //  FORM & HEADER INIT
        // =====================================================================

        private void InitializeForm()
        {
            this.Text = "Project Aura - Windows Automator";
            this.Size = new Size(1100, 950);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.None;
            this.MaximizeBox = false;
            this.BackColor = Color.FromArgb(28, 28, 28);
            this.ForeColor = Color.White;
            this.RightToLeft = RightToLeft.No;
            this.RightToLeftLayout = false;
            this.AutoScaleMode = AutoScaleMode.Dpi;

            // 1px border
            this.Paint += (_, e) =>
            {
                using var p = new Pen(Color.FromArgb(50, 50, 50), 2);
                e.Graphics.DrawRectangle(p, 0, 0, Width - 1, Height - 1);
            };

            // Timer infrastructure
            processStopwatch = new Stopwatch();
            elapsedTimer = new System.Windows.Forms.Timer { Interval = 1000 };
            elapsedTimer.Tick += (_, _) =>
            {
                if (processStopwatch.IsRunning)
                {
                    TimeSpan ts = processStopwatch.Elapsed;
                    lblTimer.Text = $"⏱ {ts.Hours:D2}:{ts.Minutes:D2}:{ts.Seconds:D2}";
                }
            };
        }

        private void InitializeHeader()
        {
            headerPanel = new Panel
            {
                Dock = DockStyle.Top,
                Height = 220,
                BackColor = Color.FromArgb(20, 20, 20),
                Cursor = Cursors.Default
            };
            headerPanel.Paint += HeaderPanel_Paint;
            headerPanel.MouseDown += (_, e) =>
            {
                if (e.Button == MouseButtons.Left)
                {
                    ReleaseCapture();
                    SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
                }
            };

            var lblClose = new Label
            {
                Text = "✕",
                Font = new Font("Segoe UI", 12, FontStyle.Bold),
                Location = new Point(1060, 10),
                AutoSize = true,
                ForeColor = Color.DarkGray,
                Cursor = Cursors.Hand,
                BackColor = Color.Transparent
            };
            lblClose.MouseEnter += (_, _) => lblClose.ForeColor = Color.LightCoral;
            lblClose.MouseLeave += (_, _) => lblClose.ForeColor = Color.DarkGray;
            lblClose.Click += (_, _) => Application.Exit();
            headerPanel.Controls.Add(lblClose);

            // === Banner Image ===
            try
            {
                var pbBanner = new PictureBox
                {
                    Dock = DockStyle.Fill,
                    Image = LoadImageFromResource("TrumpHeader.jpg"),
                    SizeMode = PictureBoxSizeMode.Zoom,
                    BackColor = Color.Transparent
                };
                pbBanner.MouseDown += (_, e) =>
                {
                    if (e.Button == MouseButtons.Left)
                    {
                        ReleaseCapture();
                        SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
                    }
                };
                headerPanel.Controls.Add(pbBanner);
            }
            catch { /* Ignore if image fails to load */ }

            this.Controls.Add(headerPanel);
        }

        private Image LoadImageFromResource(string name)
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            string resourceName = "WinAutomator." + name;
            using var stream = assembly.GetManifestResourceStream(resourceName);
            return stream != null ? Image.FromStream(stream) : null;
        }

        private void HeaderPanel_Paint(object sender, PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            using Font fontTitle = new Font("Segoe UI", 18, FontStyle.Bold);
            using Font fontCredit = new Font("Segoe UI", 11, FontStyle.Regular);

            SizeF titleSize = e.Graphics.MeasureString("Project Aura", fontTitle);
            float titleX = (headerPanel.Width - titleSize.Width) / 2f;
            e.Graphics.DrawString("Project Aura", fontTitle, Brushes.Cyan, new PointF(titleX, 8));

            StringFormat sfCenter = new StringFormat { Alignment = StringAlignment.Center };
            RectangleF rect = new RectangleF(0, 48, headerPanel.Width, 30);
            e.Graphics.DrawString("Built by Yakir Lavi", fontCredit, Brushes.LightGray, rect, sfCenter);
        }

        // =====================================================================
        //  INPUT SCREEN (DPI-Responsive with TableLayoutPanel)
        // =====================================================================

        private void InitializeInputScreen()
        {
            inputPanel = new Panel
            {
                Location = new Point(0, 220),
                Size = new Size(1100, 730),
                BackColor = Color.Transparent,
                Visible = false
            };

            Font fLabel = new Font("Segoe UI", 10, FontStyle.Bold);
            Font fBox = new Font("Segoe UI", 10);

            // --- TableLayoutPanel for input fields ---
            var table = new TableLayoutPanel
            {
                Location = new Point(250, 80),
                Size = new Size(700, 260),
                ColumnCount = 2,
                RowCount = 4,
                BackColor = Color.Transparent,
                CellBorderStyle = TableLayoutPanelCellBorderStyle.None
            };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 480));
            table.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            for (int i = 0; i < 4; i++)
                table.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));

            // Row 0: Tech Name
            txtTech = CreateTextBox(fBox);
            table.Controls.Add(txtTech, 0, 0);
            table.Controls.Add(CreateLabel("שם הטכנאי:", fLabel), 1, 0);

            // Row 1: Manufacturer
            cmbManufacturer = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = fBox,
                BackColor = Color.FromArgb(40, 40, 42),
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat,
                Dock = DockStyle.Fill,
                RightToLeft = RightToLeft.No
            };
            cmbManufacturer.Items.AddRange(AppConfig.Current.Manufacturers);
            cmbManufacturer.SelectedIndex = 0;
            table.Controls.Add(cmbManufacturer, 0, 1);
            table.Controls.Add(CreateLabel("יצרן מחשב:", fLabel), 1, 1);

            // Row 2: Serial Number
            txtSerial = CreateTextBox(fBox);
            table.Controls.Add(txtSerial, 0, 2);
            table.Controls.Add(CreateLabel("מספר סידורי:", fLabel), 1, 2);

            // Row 3: CPU Generation
            txtCpu = CreateTextBox(fBox);
            txtCpu.Text = AutomationLogic.GetCpuGeneration();
            table.Controls.Add(txtCpu, 0, 3);
            table.Controls.Add(CreateLabel("דור מעבד:", fLabel), 1, 3);

            inputPanel.Controls.Add(table);

            // --- Skip Hostname Checkbox ---
            chkSkipHostname = new CheckBox
            {
                Text = "דלג על שינוי שם המחשב (מצב בדיקה)",
                Location = new Point(350, 370),
                AutoSize = true,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.FromArgb(255, 170, 0),
                RightToLeft = RightToLeft.Yes
            };
            inputPanel.Controls.Add(chkSkipHostname);
            
            inputPanel.Controls.Add(chkSkipHostname);

            // --- Action Buttons ---
            btnFullAuto = CreateActionButton("אוטומציה מלאה", Color.FromArgb(32, 160, 100),
                Color.FromArgb(40, 190, 120), new Point(600, 480));
            btnFullAuto.Click += (_, _) => StartFresh(true);
            inputPanel.Controls.Add(btnFullAuto);

            btnSemiAuto = CreateActionButton("חצי אוטומטי", Color.FromArgb(40, 120, 180),
                Color.FromArgb(60, 140, 200), new Point(320, 480));
            btnSemiAuto.Click += (_, _) => StartFresh(false);
            inputPanel.Controls.Add(btnSemiAuto);

            // --- Jump to QA Button ---
            if (File.Exists(@"C:\WinAutomator_Completed.tag"))
            {
                var btnJumpQA = CreateActionButton("קפוץ ישירות לבדיקות חומרה", Color.FromArgb(200, 100, 30),
                    Color.FromArgb(220, 120, 50), new Point(420, 570));
                btnJumpQA.Click += (_, _) => JumpToQA();
                inputPanel.Controls.Add(btnJumpQA);
            }

            this.Controls.Add(inputPanel);
        }

        private void JumpToQA()
        {
            context.TechName = txtTech.Text.Trim() == "" ? "טכנאי מעבדה" : txtTech.Text.Trim();
            context.SerialNum = txtSerial.Text.Trim() == "" ? "Unknown" : txtSerial.Text.Trim();
            context.CpuGen = txtCpu.Text.Trim() == "" ? "Unknown" : txtCpu.Text.Trim();
            context.SelectedManufacturer = cmbManufacturer.SelectedItem?.ToString() ?? "אחר";
            
            ShowAutomationScreen();
            StartGlobalTimer();
            AppendLog("מתחיל אבחון חומרה (QA) ישירות...");
            
            // Phase 3 is index 2
            currentPhase = 3;
            RunCurrentPhase();
        }

        private static TextBox CreateTextBox(Font f) => new TextBox
        {
            Font = f,
            BackColor = Color.FromArgb(40, 40, 42),
            ForeColor = Color.White,
            BorderStyle = BorderStyle.FixedSingle,
            RightToLeft = RightToLeft.No,
            Dock = DockStyle.Fill
        };

        private static Label CreateLabel(string text, Font f) => new Label
        {
            Text = text,
            Font = f,
            AutoSize = true,
            ForeColor = Color.White,
            RightToLeft = RightToLeft.Yes,
            Anchor = AnchorStyles.Right,
            TextAlign = ContentAlignment.MiddleRight,
            Padding = new Padding(0, 10, 0, 0)
        };

        private static Button CreateActionButton(string text, Color bg, Color hoverBg, Point location)
        {
            var btn = new Button
            {
                Text = text,
                Location = location,
                Width = 250,
                Height = 55,
                FlatStyle = FlatStyle.Flat,
                BackColor = bg,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            btn.FlatAppearance.BorderSize = 0;
            btn.MouseEnter += (_, _) => btn.BackColor = hoverBg;
            btn.MouseLeave += (_, _) => btn.BackColor = bg;
            return btn;
        }

        // =====================================================================
        //  AUTOMATION SCREEN (Stepper + Timer + Log)
        // =====================================================================

        private void InitializeAutomationScreen()
        {
            automationPanel = new Panel
            {
                Location = new Point(0, 220),
                Size = new Size(1100, 730),
                BackColor = Color.Transparent,
                Visible = false
            };

            // Manufacturer info label
            lblManInfo = new Label
            {
                Location = new Point(10, 5),
                Size = new Size(1070, 22),
                TextAlign = ContentAlignment.TopCenter,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.Yellow,
                Visible = false
            };
            automationPanel.Controls.Add(lblManInfo);

            // Status label
            lblStatus = new Label
            {
                Text = "ממתין לפקודה...",
                Location = new Point(10, 28),
                Size = new Size(1070, 22),
                TextAlign = ContentAlignment.TopCenter,
                Font = new Font("Segoe UI", 10, FontStyle.Bold),
                ForeColor = Color.Cyan
            };
            automationPanel.Controls.Add(lblStatus);

            // Stepper Control (replaces ProgressBar)
            stepperControl = new StepperControl
            {
                Location = new Point(15, 55),
                Size = new Size(1060, 360),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            automationPanel.Controls.Add(stepperControl);

            // Timer
            lblTimer = new Label
            {
                Text = "⏱ 00:00:00",
                Location = new Point(30, 420),
                Size = new Size(200, 28),
                Font = new Font("Consolas", 14, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 200, 255),
                TextAlign = ContentAlignment.MiddleLeft,
                RightToLeft = RightToLeft.No
            };
            automationPanel.Controls.Add(lblTimer);

            // Log Panel
            logPanel = new Panel
            {
                Location = new Point(15, 455),
                Size = new Size(1060, 260),
                BackColor = Color.FromArgb(18, 18, 18),
                BorderStyle = BorderStyle.FixedSingle
            };

            rtbLog = new RichTextBox
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(18, 18, 18),
                ForeColor = Color.FromArgb(180, 220, 180),
                Font = new Font("Consolas", 8),
                ReadOnly = true,
                BorderStyle = BorderStyle.None,
                ScrollBars = RichTextBoxScrollBars.Vertical,
                RightToLeft = RightToLeft.No,
                WordWrap = true
            };
            logPanel.Controls.Add(rtbLog);
            automationPanel.Controls.Add(logPanel);

            this.Controls.Add(automationPanel);
        }

        // =====================================================================
        //  SCREEN TRANSITIONS
        // =====================================================================

        private void ShowInputScreen()
        {
            inputPanel.Visible = true;
            automationPanel.Visible = false;
        }

        private void ShowAutomationScreen()
        {
            inputPanel.Visible = false;
            automationPanel.Visible = true;

            if (!string.IsNullOrEmpty(context.SelectedManufacturer))
            {
                lblManInfo.Text = $"יצרן נבחר: {context.SelectedManufacturer}";
                lblManInfo.Visible = true;
            }
        }

        // =====================================================================
        //  START & PHASE PIPELINE
        // =====================================================================

        private void StartFresh(bool fullAuto)
        {
            if (string.IsNullOrWhiteSpace(txtTech.Text) ||
                string.IsNullOrWhiteSpace(txtSerial.Text) ||
                string.IsNullOrWhiteSpace(txtCpu.Text))
            {
                MessageBox.Show("נא למלא את כל 3 השדות טרם תחילת האוטומציה.",
                    "שגיאה", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            context.TechName = txtTech.Text.Trim();
            context.SerialNum = txtSerial.Text.Trim();
            context.CpuGen = txtCpu.Text.Trim();
            context.IsFullAuto = fullAuto;
            context.SkipHostname = chkSkipHostname?.Checked ?? false;
            context.SelectedManufacturer = cmbManufacturer.SelectedItem?.ToString() ?? "Lenovo";

            ShowAutomationScreen();
            StartGlobalTimer();

            AppendLog($"התחלת תהליך ({(fullAuto ? "אוטומציה מלאה" : "חצי אוטומטי")})");
            AppendLog($"טכנאי: {context.TechName} | סיריאלי: {context.SerialNum} | דור CPU: {context.CpuGen} | יצרן: {context.SelectedManufacturer}");

            RunCurrentPhase();
        }

        private async void RunCurrentPhase()
        {
            int phaseIdx = currentPhase - 1; // Convert 1-based to 0-based
            if (phaseIdx < 0 || phaseIdx >= allPhases.Length) return;

            var phase = allPhases[phaseIdx];
            var config = AppConfig.Current;

            // Setup stepper with all phases, highlighting the current one
            stepperControl.SetPhases(allPhases, phaseIdx);
            UpdateStatus(phase.PhaseName);

            // Execute the phase
            var result = await phase.ExecuteAsync(
                context,
                AppendLog,
                (stepIdx, status) => stepperControl.UpdateStep(phaseIdx, stepIdx, status));

            // Handle result
            switch (result)
            {
                case PhaseResult.RestartLoop:
                    UpdateStatus("כשל דווח! מקנפג חזרה ללופ ומבצע הפעלה מחדש...");
                    await Task.Delay(config.Timeouts.PhaseTransitionDelayMs);
                    AutomationLogic.SetupAutoResumeAndRestart(
                        currentPhase, context.TechName, context.SerialNum, context.CpuGen,
                        context.IsFullAuto, context.SkipHostname, context.SelectedManufacturer);
                    break;

                case PhaseResult.RestartAdvance:
                    stepperControl.MarkPhaseComplete(phaseIdx);
                    UpdateStatus($"{phase.PhaseName} הסתיימה! מכין מחשב להפעלה מחדש...");
                    await Task.Delay(config.Timeouts.PhaseTransitionDelayMs);
                    AutomationLogic.SetupAutoResumeAndRestart(
                        currentPhase + 1, context.TechName, context.SerialNum, context.CpuGen,
                        context.IsFullAuto, context.SkipHostname, context.SelectedManufacturer);
                    break;

                case PhaseResult.Success:
                    // Only Phase 3 returns Success – continue to interactive QA
                    await HandlePhase3Interactive(phaseIdx);
                    break;
            }
        }

        // =====================================================================
        //  PHASE 3: INTERACTIVE QA (UI-dependent, stays in MainForm)
        // =====================================================================

        private async Task HandlePhase3Interactive(int phaseIdx)
        {
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
            UpdateStatus("פותח את מסך האבחון / QA...");

            // === Step 5: Interactive Hardware QA Tests ===
            stepperControl.UpdateStep(phaseIdx, 5, StepStatus.Running);
            AppendLog("פותח מסך אבחון חומרה אינטראקטיבי");

            DialogResult micRes = DialogResult.None, stereoRes = DialogResult.None,
                camRes = DialogResult.None, kbRes = DialogResult.None,
                tpRes = DialogResult.None, usbRes = DialogResult.None;

            using var overlay = new Form
            {
                StartPosition = FormStartPosition.Manual,
                Location = this.Location,
                Size = this.Size,
                BackColor = Color.Black,
                Opacity = 0.6,
                FormBorderStyle = FormBorderStyle.None,
                ShowInTaskbar = false
            };
            overlay.Show(this);

            using var qaForm = new QaDiagnosticsForm();
            qaForm.ShowDialog(overlay);

            micRes = qaForm.MicResult;
            stereoRes = qaForm.SpeakerResult;
            camRes = qaForm.CameraResult;
            kbRes = qaForm.KeyboardResult;
            tpRes = qaForm.TrackpadResult;
            usbRes = qaForm.UsbResult;

            overlay.Close();

            stepperControl.UpdateStep(phaseIdx, 5, StepStatus.Completed);

            AppendLog($"מיקרופון: {(micRes == DialogResult.Yes ? "תקין ✓" : "נכשל ✗")}");
            AppendLog($"רמקולים: {(stereoRes == DialogResult.Yes ? "תקין ✓" : "נכשל ✗")}");
            AppendLog($"מצלמה: {(camRes == DialogResult.Yes ? "תקין ✓" : "נכשל ✗")}");
            AppendLog($"מקלדת: {(kbRes == DialogResult.Yes ? "תקין ✓" : "נכשל ✗")}");
            AppendLog($"משטח מגע: {(tpRes == DialogResult.Yes ? "תקין ✓" : "נכשל ✗")}");
            AppendLog($"כניסות USB: {(usbRes == DialogResult.Yes ? "תקין ✓" : (usbRes == DialogResult.Ignore ? "דולג ⏭" : "נכשל ✗"))}");

            stepperControl.MarkPhaseComplete(phaseIdx);

            // === STOP TIMER ===
            StopGlobalTimer();
            string totalTime = GetElapsedTimeString();
            context.ElapsedTime = totalTime;
            AppendLog($"═══ התהליך הושלם! זמן כולל: {totalTime} ═══");

            try { File.WriteAllText(@"C:\WinAutomator_Completed.tag", DateTime.Now.ToString()); } catch { }

            ShowResultsSummary(micRes, camRes, kbRes, tpRes, usbRes, context.SsdHealth, stereoRes, totalTime);
        }

        // =====================================================================
        //  TIMER
        // =====================================================================

        private void StartGlobalTimer()
        {
            lblTimer.Visible = true;
            if (!processStopwatch.IsRunning) processStopwatch.Start();
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

        // =====================================================================
        //  LOG & STATUS
        // =====================================================================

        private void AppendLog(string message)
        {
            if (InvokeRequired) { Invoke(new Action(() => AppendLog(message))); return; }

            string entry = $"[{DateTime.Now:HH:mm:ss}] {message}";
            logEntries.Add(entry);
            rtbLog.AppendText(entry + "\n");
            rtbLog.ScrollToCaret();
        }

        private void UpdateStatus(string message)
        {
            if (InvokeRequired) { Invoke(new Action(() => UpdateStatus(message))); return; }
            lblStatus.Text = message;
            AppendLog(message);
        }

        // =====================================================================
        //  RESULTS SUMMARY
        // =====================================================================

        private void ShowResultsSummary(DialogResult mic, DialogResult cam, DialogResult kb,
            DialogResult tp, DialogResult usb, string ssd, DialogResult stereo, string totalTime)
        {
            if (InvokeRequired)
            {
                Invoke(new Action(() => ShowResultsSummary(mic, cam, kb, tp, usb, ssd, stereo, totalTime)));
                return;
            }

            // Hide automation panel
            automationPanel.Visible = false;

            // Resize form for results
            this.Height = 750;

            Panel resPanel = new Panel
            {
                Location = new Point(150, 130),
                Size = new Size(800, 560),
                BackColor = Color.FromArgb(40, 40, 40),
                BorderStyle = BorderStyle.FixedSingle
            };

            Label title = new Label
            {
                Text = "סיכום בדיקות שבוצעו",
                Dock = DockStyle.Top,
                Height = 35,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Segoe UI", 12, FontStyle.Bold)
            };
            resPanel.Controls.Add(title);

            string[] names = { "מיקרופון", "מצלמה", "מקלדת", "משטח מגע", "שקעי USB", "סטריאו רמקולים" };
            DialogResult[] results = { mic, cam, kb, tp, usb, stereo };

            for (int i = 0; i < names.Length; i++)
            {
                Label lblName = new Label
                {
                    Text = names[i],
                    Location = new Point(550, 50 + i * 45),
                    AutoSize = true,
                    Font = new Font("Segoe UI", 10)
                };

                string resTxt = "נכשל ✗"; Color resCol = Color.LightCoral;
                if (results[i] == DialogResult.Yes) { resTxt = "תקין ✓"; resCol = Color.LimeGreen; }
                else if (results[i] == DialogResult.Ignore) { resTxt = "דולג ⏭"; resCol = Color.Goldenrod; }

                Label lblRes = new Label
                {
                    Text = resTxt,
                    ForeColor = resCol,
                    Location = new Point(100, 50 + i * 45),
                    AutoSize = true,
                    Font = new Font("Segoe UI", 10, FontStyle.Bold)
                };
                resPanel.Controls.Add(lblName);
                resPanel.Controls.Add(lblRes);
            }

            // SSD Info
            resPanel.Controls.Add(new Label
            {
                Text = "בריאות דיסק:",
                Location = new Point(550, 50 + names.Length * 45),
                AutoSize = true,
                Font = new Font("Segoe UI", 10)
            });
            resPanel.Controls.Add(new Label
            {
                Text = ssd,
                Location = new Point(60, 50 + names.Length * 45),
                Width = 400,
                ForeColor = Color.Cyan,
                Font = new Font("Segoe UI", 9, FontStyle.Bold)
            });

            // Total Time
            int timeY = 50 + (names.Length + 1) * 45;
            resPanel.Controls.Add(new Label
            {
                Text = "⏱ זמן כולל:",
                Location = new Point(300, timeY),
                AutoSize = true,
                Font = new Font("Segoe UI", 10)
            });
            resPanel.Controls.Add(new Label
            {
                Text = totalTime,
                Location = new Point(50, timeY),
                AutoSize = true,
                ForeColor = Color.FromArgb(0, 200, 255),
                Font = new Font("Consolas", 12, FontStyle.Bold),
                RightToLeft = RightToLeft.No
            });

            // Show Log Button
            Button btnShowLog = new Button
            {
                Text = "📋 הצג לוג מלא",
                Location = new Point(80, 490),
                Width = 240,
                Height = 44,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(55, 55, 55),
                ForeColor = Color.LightGray,
                Font = new Font("Segoe UI", 9),
                Cursor = Cursors.Hand
            };
            btnShowLog.FlatAppearance.BorderColor = Color.FromArgb(80, 80, 80);
            btnShowLog.Click += (_, _) =>
            {
                Form logForm = new Form
                {
                    Text = "Detail Log - Project Aura",
                    Size = new Size(700, 500),
                    StartPosition = FormStartPosition.CenterParent,
                    BackColor = Color.FromArgb(18, 18, 18),
                    ForeColor = Color.White
                };
                RichTextBox fullLog = new RichTextBox
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

            // Finish Button
            Button btnFinish = new Button
            {
                Text = "סיום וסגירה",
                Location = new Point(460, 490),
                Width = 240,
                Height = 44,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.SeaGreen,
                Cursor = Cursors.Hand
            };
            btnFinish.Click += (_, _) =>
            {
                var rData = new ReportData
                {
                    TimeStamp = DateTime.Now.ToString("dd/MM/yyyy HH:mm"),
                    TechId = context.TechName,
                    SerialNum = context.SerialNum,
                    CpuGen = context.CpuGen,
                    Manufacturer = context.SelectedManufacturer,
                    CpuName = AutomationLogic.GetCpuName(),
                    RamSize = AutomationLogic.GetRamSizeGB() + " GB",
                    SsdSize = AutomationLogic.GetTotalDiskSizeGB() + " GB",
                    MicResult = mic,
                    SpeakerResult = stereo,
                    CameraResult = cam,
                    KeyboardResult = kb,
                    TrackpadResult = tp,
                    UsbResult = usb
                };
                ReportGenerator.GenerateAutomationReport(rData);
                MessageBox.Show("הדוח הופק בהצלחה ונשמר בשולחן העבודה!", "דוח הופק", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Application.Exit();
            };
            resPanel.Controls.Add(btnFinish);

            this.Controls.Add(resPanel);
        }
    }
}
