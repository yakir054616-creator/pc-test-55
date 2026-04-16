using System;
using System.Threading;
using System.Threading.Tasks;
using System.Collections.Generic;

namespace WinAutomator.Phases
{
    /// <summary>
    /// Status of a single sub-step within a phase.
    /// </summary>
    public enum StepStatus
    {
        Pending,
        Running,
        Completed,
        Failed,
        Skipped
    }

    /// <summary>
    /// Describes a named sub-step for UI display.
    /// </summary>
    public class StepDefinition
    {
        public string Name { get; set; }
        public StepStatus Status { get; set; } = StepStatus.Pending;

        public StepDefinition(string name) { Name = name; }
    }

    /// <summary>
    /// Result of executing a phase.
    /// </summary>
    public enum PhaseResult
    {
        /// <summary>Phase completed, advance to next phase.</summary>
        Success,
        /// <summary>Phase needs to loop (e.g. DISM failure in semi-auto). Restart and repeat.</summary>
        RestartLoop,
        /// <summary>Phase ended; system will restart and resume at next phase.</summary>
        RestartAdvance
    }

    /// <summary>
    /// Shared context passed between all phases.
    /// </summary>
    public class AutomationContext
    {
        public string TechName { get; set; } = "";
        public string SerialNum { get; set; } = "";
        public string CpuGen { get; set; } = "";
        public bool IsFullAuto { get; set; } = true;
        public bool SkipHostname { get; set; }
        public string SelectedManufacturer { get; set; } = "Lenovo";

        // Results collected during Phase 3 for the final summary
        public string SsdHealth { get; set; } = "";
        public string ElapsedTime { get; set; } = "";
    }

    /// <summary>
    /// Contract for automation phases. Each phase encapsulates its own 
    /// logic independently from the UI, enabling easy reordering, 
    /// insertion, or removal of phases.
    /// </summary>
    public interface IAutomationPhase
    {
        /// <summary>Display name for the phase (e.g., "פאזה 1: הכנת דיסק").</summary>
        string PhaseName { get; }

        /// <summary>Short description shown in the stepper.</summary>
        string Description { get; }

        /// <summary>Ordered list of sub-steps for granular UI feedback.</summary>
        List<StepDefinition> Steps { get; }

        /// <summary>
        /// Executes the phase logic.
        /// </summary>
        /// <param name="ctx">Shared data context across all phases.</param>
        /// <param name="log">Callback to append a timestamped log entry.</param>
        /// <param name="onStepChanged">Callback: (stepIndex, newStatus) – drives the Stepper UI.</param>
        /// <param name="ct">Cancellation token for graceful abort.</param>
        /// <returns>Result indicating what should happen next.</returns>
        Task<PhaseResult> ExecuteAsync(
            AutomationContext ctx,
            Action<string> log,
            Action<int, StepStatus> onStepChanged,
            CancellationToken ct = default);
    }
}
