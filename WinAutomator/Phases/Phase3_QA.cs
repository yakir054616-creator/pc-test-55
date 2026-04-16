using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace WinAutomator.Phases
{
    /// <summary>
    /// Phase 3: Hardware diagnostics, cleanup, battery report, QA.
    /// The background automation steps only. UI-driven QA tests (mic, camera, etc.)
    /// are triggered by MainForm after this phase signals a callback.
    /// </summary>
    public class Phase3_QA : IAutomationPhase
    {
        private void ShowDangerDialog(string title, string message)
        {
            var mainForm = System.Windows.Forms.Application.OpenForms[0];
            if (mainForm != null && mainForm.InvokeRequired)
            {
                mainForm.Invoke(new Action(() => ShowDangerDialogUI(title, message, mainForm)));
                return;
            }
            ShowDangerDialogUI(title, message, mainForm);
        }

        private void ShowDangerDialogUI(string title, string message, System.Windows.Forms.Form owner)
        {
            using var f = new System.Windows.Forms.Form
            {
                Text = title,
                Size = new System.Drawing.Size(800, 400),
                StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen,
                BackColor = System.Drawing.Color.DarkRed,
                ForeColor = System.Drawing.Color.White,
                FormBorderStyle = System.Windows.Forms.FormBorderStyle.None,
                TopMost = true
            };
            var lbl = new System.Windows.Forms.Label
            {
                Text = message,
                Font = new System.Drawing.Font("Segoe UI", 24, System.Drawing.FontStyle.Bold),
                Dock = System.Windows.Forms.DockStyle.Fill,
                TextAlign = System.Drawing.ContentAlignment.MiddleCenter,
                RightToLeft = System.Windows.Forms.RightToLeft.Yes
            };
            var btn = new System.Windows.Forms.Button
            {
                Text = "הבנתי, המערכת תקולה ואשייך אותה לתיקון. המשך הלאה",
                Font = new System.Drawing.Font("Segoe UI", 16, System.Drawing.FontStyle.Bold),
                Dock = System.Windows.Forms.DockStyle.Bottom,
                Height = 80,
                BackColor = System.Drawing.Color.Black,
                ForeColor = System.Drawing.Color.White,
                FlatStyle = System.Windows.Forms.FlatStyle.Flat
            };
            btn.Click += (s, e) => f.Close();
            f.Controls.Add(lbl);
            f.Controls.Add(btn);
            f.ShowDialog(owner);
        }

        public string PhaseName => "פאזה 3: אבחון חומרה ובדיקות QA";
        public string Description => "אופטימיזציה, ניקיון, בדיקות";

        public List<StepDefinition> Steps { get; } = new()
        {
            new StepDefinition("אופטימיזציית כונן (TRIM/Defrag)"),
            new StepDefinition("בדיקת בריאות דיסק (S.M.A.R.T)"),
            new StepDefinition("ניקוי מערכת מקיף"),
            new StepDefinition("ניקוי רשומות Auto-Resume"),
            new StepDefinition("דוח סוללה ושידור API"),
            new StepDefinition("בדיקות חומרה אינטראקטיביות")
        };

        public async Task<PhaseResult> ExecuteAsync(
            AutomationContext ctx,
            Action<string> log,
            Action<int, StepStatus> onStepChanged,
            CancellationToken ct)
        {
            log("═══ פאזה 3: אבחון חומרה ובדיקות QA ═══");

            // Step 0: Optimize Drive
            onStepChanged(0, StepStatus.Running);
            log("מבצע אופטימיזציה לכונן (TRIM/Defrag)...");
            string defragOutput = await Task.Run(
                () => AutomationLogic.OptimizeDrive(msg => log(msg)), ct);
            log("תוצאת Defrag/Optimize:\n" + defragOutput.Trim());
            onStepChanged(0, StepStatus.Completed);

            // Step 1: SSD Health
            onStepChanged(1, StepStatus.Running);
            log("בודק את בריאות הדיסק (S.M.A.R.T)...");
            ctx.SsdHealth = await Task.Run(() => AutomationLogic.GetSsdHealth(), ct);
            log($"בריאות דיסק: {ctx.SsdHealth}");

            if (ctx.SsdHealth.Contains("Warning") || ctx.SsdHealth.Contains("אזהרה") || ctx.SsdHealth.Contains("Error"))
            {
                log("⚠ אזהרת S.M.A.R.T התקבלה מהבקר - כונן פסול!");
                ShowDangerDialog("סכנת קריסת כונן (S.M.A.R.T)", "מערכת הבקרה מדווחת כי הכונן (SSD/HDD) תקול פיזית\nויש להחליפו באופן מיידי!");
            }

            onStepChanged(1, StepStatus.Completed);

            // Step 2: System Cleanup
            onStepChanged(2, StepStatus.Running);
            log("מבצע ניקוי סופי למערכת (Temp / Cache / Updates)...");
            await Task.Run(() => AutomationLogic.PerformSystemCleanup(msg => log(msg)), ct);
            log("ניקוי מערכת הושלם");
            onStepChanged(2, StepStatus.Completed);

            // Step 3: Cleanup Auto-Resume
            onStepChanged(3, StepStatus.Running);
            log("מנקה רישומי זיכרון שיירים (Cleanup)...");
            AutomationLogic.CleanupAutoResume();
            log("רשומות Auto-Resume נוקו");
            onStepChanged(3, StepStatus.Completed);

            // Step 4: Battery Report
            onStepChanged(4, StepStatus.Running);
            log("מפיק דו\"ח בריאות סוללה ומשדר ל-API...");
            double batteryHealth = await AutomationLogic.PerformBatteryReportAndApi(
                ctx.TechName, ctx.SerialNum, ctx.CpuGen, msg => log(msg));
            log("דוח סוללה שוגר ל-API");

            if (batteryHealth > 0 && batteryHealth < 60)
            {
                log($"⚠ אזהרת סוללה: תקינות נמוכה מ-60% ({batteryHealth}%).");
                ShowDangerDialog("סוללה תקולה", $"הסוללה הגיעה למצב שחיקה קריטי ({batteryHealth}%) חיי סוללה.\nיש להחליפה בחדשה טרם מסירה ללקוח!");
            }

            onStepChanged(4, StepStatus.Completed);

            // Victory Sound
            log("♫ משמיע צליל סיום התקנות...");
            await Task.Run(() => AutomationLogic.PlayVictorySound(), ct);

            // Steps 5 & 6 (QA interactive tests + questionnaire) are handled by MainForm
            // because they require UI thread access for modal dialogs.
            // We return Success to signal the main form to continue with the UI steps.

            return PhaseResult.Success;
        }
    }
}
