using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Windows.Forms;
using WinAutomator.Phases;

namespace WinAutomator
{
    /// <summary>
    /// A custom vertical timeline / stepper control that replaces the basic ProgressBar.
    /// Displays all phases and their sub-steps with real-time status updates and 
    /// a pulsing animation for the currently running step.
    /// </summary>
    public class StepperControl : Panel
    {
        // === Data Model ===
        private readonly List<PhaseDisplay> _phases = new();
        private int _activePhaseIndex = -1;

        // === Animation ===
        private readonly System.Windows.Forms.Timer _pulseTimer;
        private float _pulseValue = 0.3f;
        private bool _pulseUp = true;

        // === Layout Constants ===
        private const int LEFT_MARGIN = 30;
        private const int TEXT_LEFT = 55;
        private const int PHASE_RADIUS = 8;
        private const int STEP_RADIUS = 5;
        private const int PHASE_ROW_H = 36;
        private const int STEP_ROW_H = 26;
        private const int PHASE_GAP = 8;

        // === Colors ===
        private static readonly Color ColPending   = Color.FromArgb(70, 70, 70);
        private static readonly Color ColRunning   = Color.FromArgb(0, 200, 255);
        private static readonly Color ColCompleted = Color.FromArgb(80, 210, 80);
        private static readonly Color ColFailed    = Color.FromArgb(230, 80, 80);
        private static readonly Color ColSkipped   = Color.FromArgb(210, 190, 50);

        public class PhaseDisplay
        {
            public string Name { get; set; } = "";
            public bool IsExpanded { get; set; }
            public StepStatus OverallStatus { get; set; } = StepStatus.Pending;
            public List<StepDisplay> Steps { get; set; } = new();
        }

        public class StepDisplay
        {
            public string Name { get; set; } = "";
            public StepStatus Status { get; set; } = StepStatus.Pending;
        }

        public StepperControl()
        {
            this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint, true);
            this.AutoScroll = true;
            this.BackColor = Color.FromArgb(25, 25, 25);

            _pulseTimer = new System.Windows.Forms.Timer { Interval = 50 };
            _pulseTimer.Tick += (_, _) =>
            {
                _pulseValue += _pulseUp ? 0.04f : -0.04f;
                if (_pulseValue >= 1f) { _pulseValue = 1f; _pulseUp = false; }
                if (_pulseValue <= 0.2f) { _pulseValue = 0.2f; _pulseUp = true; }
                Invalidate();
            };
            _pulseTimer.Start();
        }

        /// <summary>
        /// Initializes the stepper display with all phases.
        /// Past phases (before activePhaseIndex) are shown collapsed and completed.
        /// The current phase is expanded with live sub-steps.
        /// Future phases are collapsed and pending.
        /// </summary>
        public void SetPhases(IAutomationPhase[] allPhases, int activePhaseIndex)
        {
            _phases.Clear();
            _activePhaseIndex = activePhaseIndex;

            for (int i = 0; i < allPhases.Length; i++)
            {
                var src = allPhases[i];
                var pd = new PhaseDisplay
                {
                    Name = src.PhaseName,
                    IsExpanded = (i == activePhaseIndex),
                    OverallStatus = i < activePhaseIndex ? StepStatus.Completed
                                  : i == activePhaseIndex ? StepStatus.Running
                                  : StepStatus.Pending
                };

                foreach (var step in src.Steps)
                {
                    pd.Steps.Add(new StepDisplay
                    {
                        Name = step.Name,
                        Status = i < activePhaseIndex ? StepStatus.Completed : StepStatus.Pending
                    });
                }

                _phases.Add(pd);
            }

            Invalidate();
        }

        /// <summary>
        /// Updates a specific sub-step's status. Thread-safe.
        /// </summary>
        public void UpdateStep(int phaseIndex, int stepIndex, StepStatus status)
        {
            if (InvokeRequired) { Invoke(new Action(() => UpdateStep(phaseIndex, stepIndex, status))); return; }

            if (phaseIndex < 0 || phaseIndex >= _phases.Count) return;
            var phase = _phases[phaseIndex];
            if (stepIndex < 0 || stepIndex >= phase.Steps.Count) return;

            phase.Steps[stepIndex].Status = status;
            RecalcPhaseStatus(phase);
            Invalidate();
        }

        /// <summary>
        /// Forces a phase to show as completed (all sub-steps green).
        /// </summary>
        public void MarkPhaseComplete(int phaseIndex)
        {
            if (InvokeRequired) { Invoke(new Action(() => MarkPhaseComplete(phaseIndex))); return; }

            if (phaseIndex < 0 || phaseIndex >= _phases.Count) return;
            var phase = _phases[phaseIndex];
            phase.OverallStatus = StepStatus.Completed;
            foreach (var s in phase.Steps)
                if (s.Status != StepStatus.Failed && s.Status != StepStatus.Skipped)
                    s.Status = StepStatus.Completed;
            Invalidate();
        }

        private static void RecalcPhaseStatus(PhaseDisplay phase)
        {
            bool allDone = true;
            bool anyFailed = false;
            foreach (var s in phase.Steps)
            {
                if (s.Status is StepStatus.Running or StepStatus.Pending) allDone = false;
                if (s.Status == StepStatus.Failed) anyFailed = true;
            }
            if (allDone)
                phase.OverallStatus = anyFailed ? StepStatus.Failed : StepStatus.Completed;
        }

        // ============================ PAINTING ============================

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            // Apply scroll offset
            g.TranslateTransform(AutoScrollPosition.X, AutoScrollPosition.Y);

            int y = 10;
            int contentWidth = this.ClientSize.Width - 10;

            using var fPhase = new Font("Segoe UI", 10.5f, FontStyle.Bold);
            using var fStep = new Font("Segoe UI", 9f);
            using var fIcon = new Font("Segoe UI", 9f, FontStyle.Bold);
            using var linePen = new Pen(Color.FromArgb(45, 45, 50), 2);

            for (int pi = 0; pi < _phases.Count; pi++)
            {
                var phase = _phases[pi];
                int circleY = y + PHASE_ROW_H / 2;

                // --- Vertical line from previous phase ---
                if (pi > 0) g.DrawLine(linePen, LEFT_MARGIN, y - PHASE_GAP, LEFT_MARGIN, circleY - PHASE_RADIUS);

                // --- Phase circle ---
                Color phaseCol = StatusColor(phase.OverallStatus);
                if (phase.OverallStatus == StepStatus.Running)
                    phaseCol = Color.FromArgb((int)(100 + 155 * _pulseValue), phaseCol);

                using (var brush = new SolidBrush(phaseCol))
                    g.FillEllipse(brush, LEFT_MARGIN - PHASE_RADIUS, circleY - PHASE_RADIUS, PHASE_RADIUS * 2, PHASE_RADIUS * 2);

                // Glow ring for running
                if (phase.OverallStatus == StepStatus.Running)
                {
                    using var glowPen = new Pen(Color.FromArgb((int)(50 * _pulseValue), ColRunning), 3);
                    g.DrawEllipse(glowPen, LEFT_MARGIN - PHASE_RADIUS - 4, circleY - PHASE_RADIUS - 4,
                        (PHASE_RADIUS + 4) * 2, (PHASE_RADIUS + 4) * 2);
                }

                // --- Phase text ---
                Color textCol = phase.OverallStatus == StepStatus.Pending ? Color.FromArgb(100, 100, 100) : Color.White;
                using (var tb = new SolidBrush(textCol))
                    g.DrawString(phase.Name, fPhase, tb, TEXT_LEFT, y + 8);

                // Status icon
                string icon = StatusIcon(phase.OverallStatus);
                if (!string.IsNullOrEmpty(icon))
                {
                    using var ib = new SolidBrush(phaseCol);
                    var isz = g.MeasureString(icon, fIcon);
                    g.DrawString(icon, fIcon, ib, contentWidth - isz.Width - 15, y + 9);
                }

                y += PHASE_ROW_H;

                // --- Sub-steps (only if expanded) ---
                if (phase.IsExpanded)
                {
                    for (int si = 0; si < phase.Steps.Count; si++)
                    {
                        var step = phase.Steps[si];
                        int sCircleY = y + STEP_ROW_H / 2;

                        // Vertical line
                        g.DrawLine(linePen, LEFT_MARGIN, y - 4, LEFT_MARGIN, sCircleY - STEP_RADIUS);

                        // Step circle
                        Color sCol = StatusColor(step.Status);
                        if (step.Status == StepStatus.Running)
                            sCol = Color.FromArgb((int)(100 + 155 * _pulseValue), ColRunning);

                        using (var sb = new SolidBrush(sCol))
                            g.FillEllipse(sb, LEFT_MARGIN - STEP_RADIUS, sCircleY - STEP_RADIUS, STEP_RADIUS * 2, STEP_RADIUS * 2);

                        // Step text
                        Color sTextCol = step.Status == StepStatus.Pending ? Color.FromArgb(90, 90, 90) : Color.FromArgb(195, 195, 195);
                        using (var stb = new SolidBrush(sTextCol))
                            g.DrawString(step.Name, fStep, stb, TEXT_LEFT, y + 4);

                        // Step status icon
                        string sIcon = StatusIcon(step.Status);
                        if (!string.IsNullOrEmpty(sIcon))
                        {
                            using var sib = new SolidBrush(sCol);
                            var sz = g.MeasureString(sIcon, fIcon);
                            g.DrawString(sIcon, fIcon, sib, contentWidth - sz.Width - 15, y + 4);
                        }

                        // Line after step
                        if (si < phase.Steps.Count - 1)
                            g.DrawLine(linePen, LEFT_MARGIN, sCircleY + STEP_RADIUS, LEFT_MARGIN, y + STEP_ROW_H);

                        y += STEP_ROW_H;
                    }
                }

                // Line to next phase
                if (pi < _phases.Count - 1)
                {
                    g.DrawLine(linePen, LEFT_MARGIN, y - 3, LEFT_MARGIN, y + PHASE_GAP);
                    y += PHASE_GAP;
                }
            }

            // Update scrollable area
            int requiredHeight = y + 10;
            if (AutoScrollMinSize.Height != requiredHeight)
                AutoScrollMinSize = new Size(0, requiredHeight);
        }

        // ============================ HELPERS ============================

        private static Color StatusColor(StepStatus s) => s switch
        {
            StepStatus.Running   => ColRunning,
            StepStatus.Completed => ColCompleted,
            StepStatus.Failed    => ColFailed,
            StepStatus.Skipped   => ColSkipped,
            _                    => ColPending
        };

        private static string StatusIcon(StepStatus s) => s switch
        {
            StepStatus.Completed => "✓",
            StepStatus.Failed    => "✗",
            StepStatus.Running   => "⟳",
            StepStatus.Skipped   => "⏭",
            _                    => ""
        };

        protected override void Dispose(bool disposing)
        {
            if (disposing) { _pulseTimer?.Stop(); _pulseTimer?.Dispose(); }
            base.Dispose(disposing);
        }
    }
}
