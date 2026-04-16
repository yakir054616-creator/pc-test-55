using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace WinAutomator.Phases
{
    /// <summary>
    /// Phase 2: Hostname, licensing, OEM drivers, and Windows Updates.
    /// Encapsulates ChangeHostname, ActivateWindows/Office, RunOemUpdates, RunWindowsUpdates.
    /// </summary>
    public class Phase2_Updates : IAutomationPhase
    {
        public string PhaseName => "פאזה 2: רשיונות, דרייברים ועדכונים";
        public string Description => "אקטיבציות, OEM, Windows Update";

        public List<StepDefinition> Steps { get; } = new()
        {
            new StepDefinition("שינוי שם מחשב (Hostname)"),
            new StepDefinition("אקטיבציית Windows 11"),
            new StepDefinition("אקטיבציית Office 2021"),
            new StepDefinition("עדכוני יצרן (OEM)"),
            new StepDefinition("השלמות חומרה אוניברסליות (Fallback)"),
            new StepDefinition("עדכוני Windows"),
            new StepDefinition("המתנה לחילוץ חבילות")
        };

        public async Task<PhaseResult> ExecuteAsync(
            AutomationContext ctx,
            Action<string> log,
            Action<int, StepStatus> onStepChanged,
            CancellationToken ct)
        {
            log("═══ פאזה 2: רשיונות, דרייברים ועדכונים ═══");
            var config = AppConfig.Current;

            // Step 0: Hostname
            onStepChanged(0, StepStatus.Running);
            if (ctx.SkipHostname)
            {
                log("דילוג על Hostname (בקשת משתמש)");
                onStepChanged(0, StepStatus.Skipped);
            }
            else
            {
                log("מעדכן שם מחשב (Hostname)...");
                await Task.Run(() => AutomationLogic.ChangeHostname(ctx.SerialNum), ct);
                log($"Hostname שונה ל: {ctx.SerialNum}");
                onStepChanged(0, StepStatus.Completed);
            }

            // Step 1: Windows Activation
            onStepChanged(1, StepStatus.Running);
            log("מבצע אקטיבציה ל-Windows 11...");
            await Task.Run(() => AutomationLogic.ActivateWindows(), ct);
            log("אקטיבציית Windows הושלמה");
            onStepChanged(1, StepStatus.Completed);

            // Step 2: Office Activation
            onStepChanged(2, StepStatus.Running);
            log("מבצע אקטיבציה ל-Office 2021...");
            await Task.Run(() => AutomationLogic.ActivateOffice(), ct);
            log("אקטיבציית Office הושלמה");
            onStepChanged(2, StepStatus.Completed);

            // Step 3: OEM Updates
            onStepChanged(3, StepStatus.Running);
            await Task.Run(() => AutomationLogic.RunOemUpdates(
                msg => { log(msg); }, ctx.SelectedManufacturer), ct);
            log("══════════════════════════════════════");
            log($"✅  עדכוני יצרן ({ctx.SelectedManufacturer}) עודכנו בהצלחה!");
            log("══════════════════════════════════════");
            onStepChanged(3, StepStatus.Completed);
            await Task.Delay(2000, ct); // השהייה של 2 שניות כדי שהטכנאי יראה את ההודעה

            // Step 4: Universal Driver Fallback (Yellow Bang Check)
            onStepChanged(4, StepStatus.Running);
            log("בודק חוסר בדרייברים דרך מנהל ההתקנים (Device Manager)...");
            bool isMissing = false;
            await Task.Run(() => { isMissing = AutomationLogic.CheckForMissingDrivers(); }, ct);
            
            if (isMissing)
            {
                log("זוהו רכיבי חומרה חסרים (Yellow Bangs)! מתחיל תהליך גיבוי אוניברסלי.");
                await Task.Run(() => AutomationLogic.RunUniversalDriverUpdate(msg => { log(msg); }), ct);
                onStepChanged(4, StepStatus.Completed);
            }
            else
            {
                log("מערכת ההפעלה מדווחת כי מנהל ההתקנים תקין לחלוטין (0 שגיאות), דילוג על מנוע אוניברסלי.");
                onStepChanged(4, StepStatus.Skipped);
            }

            // Step 5: Windows Updates (fire & forget)
            onStepChanged(5, StepStatus.Running);
            await Task.Run(() => AutomationLogic.RunWindowsUpdates(
                msg => { log(msg); }), ct);
            log("בקשת עדכוני Windows נשלחה");
            onStepChanged(5, StepStatus.Completed);

            // Step 6: Wait for packages
            onStepChanged(6, StepStatus.Running);
            log($"ממתין {config.Timeouts.WindowsUpdateWaitMs / 1000} שניות לחילוץ חבילות...");
            await Task.Delay(config.Timeouts.WindowsUpdateWaitMs, ct);
            log("עדכוני Windows סיימו המתנה לחילוץ חבילות");
            onStepChanged(6, StepStatus.Completed);

            log("✓ פאזה 2 הושלמה בהצלחה – מכין ריסטארט לפאזה 3");
            return PhaseResult.RestartAdvance;
        }
    }
}
