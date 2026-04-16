using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WinAutomator.Phases
{
    /// <summary>
    /// Phase 1: Disk expansion + DISM system repair.
    /// Encapsulates ExtendCDrive and RunDismCommand logic.
    /// </summary>
    public class Phase1_DiskRepair : IAutomationPhase
    {
        public string PhaseName => "פאזה 1: הכנת דיסק ותיקון ליבה";
        public string Description => "הרחבת מחיצה, סריקת DISM";

        public List<StepDefinition> Steps { get; } = new()
        {
            new StepDefinition("הרחבת כונן C"),
            new StepDefinition("תיקון ליבת מערכת (DISM)")
        };

        public async Task<PhaseResult> ExecuteAsync(
            AutomationContext ctx,
            Action<string> log,
            Action<int, StepStatus> onStepChanged,
            CancellationToken ct)
        {
            log("═══ פאזה 1: הכנת דיסק ותיקון ליבה ═══");

            // Step 0: Extend C Drive
            onStepChanged(0, StepStatus.Running);
            log("סורק ומרחיב דיסק (במידה וניתן)...");
            await Task.Run(() => AutomationLogic.ExtendCDrive(), ct);
            log("הרחבת דיסק C הושלמה");
            onStepChanged(0, StepStatus.Completed);

            // Step 1: DISM
            onStepChanged(1, StepStatus.Running);
            log("מריץ פקודת DISM לשיקום מערכת. נא להמתין...");
            bool dismSuccess = await Task.Run(() => AutomationLogic.RunDismCommand(ctx.IsFullAuto), ct);
            log($"DISM הסתיים (הצלחה: {dismSuccess})");

            if (!ctx.IsFullAuto)
            {
                DialogResult res = MessageBox.Show(
                    "האם תהליך ה-DISM הסתיים בהצלחה בלי שגיאות אדומות?",
                    "בקרת DISM חציונית",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (res == DialogResult.No)
                {
                    log("⚠ כשל DISM דווח ע\"י הטכנאי – מתחיל לופ ריסטארט");
                    onStepChanged(1, StepStatus.Failed);
                    return PhaseResult.RestartLoop;
                }
            }

            onStepChanged(1, StepStatus.Completed);
            log("✓ פאזה 1 הושלמה בהצלחה – מכין ריסטארט לפאזה 2");
            return PhaseResult.RestartAdvance;
        }
    }
}
