using System;
using System.Diagnostics;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace WinAutomator
{
    /// <summary>
    /// Result of a process execution.
    /// </summary>
    public class ProcessResult
    {
        public int ExitCode { get; set; } = -1;
        public string Output { get; set; } = "";
        public string Error { get; set; } = "";
        public bool TimedOut { get; set; }
        public bool Success => ExitCode == 0 && !TimedOut;
    }

    /// <summary>
    /// Centralized, robust process runner with timeout support, cancellation,
    /// and consistent error handling. Replaces scattered Process.Start calls.
    /// </summary>
    public static class ProcessRunner
    {
        /// <summary>
        /// Runs a process silently (hidden window, no shell execute).
        /// Captures stdout if redirectOutput is true.
        /// </summary>
        public static ProcessResult RunSilent(
            string fileName,
            string arguments,
            int? timeoutMs = null,
            bool redirectOutput = false,
            CancellationToken ct = default)
        {
            int timeout = timeoutMs ?? AppConfig.Current.Timeouts.ProcessDefaultTimeoutMs;
            var result = new ProcessResult();

            try
            {
                var psi = new ProcessStartInfo(fileName, arguments)
                {
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    RedirectStandardOutput = redirectOutput,
                    RedirectStandardError = redirectOutput
                };

                using var process = Process.Start(psi);
                if (process == null)
                {
                    result.Error = $"Failed to start process: {fileName}";
                    return result;
                }

                if (redirectOutput)
                {
                    // Read output asynchronously to avoid deadlocks
                    var outputTask = process.StandardOutput.ReadToEndAsync();
                    var errorTask = process.StandardError.ReadToEndAsync();

                    bool exited = process.WaitForExit(timeout);
                    if (!exited)
                    {
                        result.TimedOut = true;
                        try { process.Kill(entireProcessTree: true); } catch { }
                        result.Error = $"Process timed out after {timeout}ms: {fileName} {arguments}";
                        return result;
                    }

                    // Ensure async reads complete after process exit
                    result.Output = outputTask.GetAwaiter().GetResult();
                    result.Error = errorTask.GetAwaiter().GetResult();
                }
                else
                {
                    bool exited = process.WaitForExit(timeout);
                    if (!exited)
                    {
                        result.TimedOut = true;
                        try { process.Kill(entireProcessTree: true); } catch { }
                        result.Error = $"Process timed out after {timeout}ms: {fileName} {arguments}";
                        return result;
                    }
                }

                result.ExitCode = process.ExitCode;
            }
            catch (Exception ex)
            {
                result.Error = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// Runs a process with a visible window (UseShellExecute = true).
        /// Cannot capture output in this mode.
        /// </summary>
        public static ProcessResult RunVisible(
            string fileName,
            string arguments,
            int? timeoutMs = null,
            CancellationToken ct = default)
        {
            int timeout = timeoutMs ?? AppConfig.Current.Timeouts.ProcessDefaultTimeoutMs;
            var result = new ProcessResult();

            try
            {
                var psi = new ProcessStartInfo(fileName, arguments)
                {
                    UseShellExecute = true,
                    CreateNoWindow = false
                };

                using var process = Process.Start(psi);
                if (process == null)
                {
                    result.Error = $"Failed to start process: {fileName}";
                    return result;
                }

                bool exited = process.WaitForExit(timeout);
                if (!exited)
                {
                    result.TimedOut = true;
                    try { process.Kill(entireProcessTree: true); } catch { }
                    result.Error = $"Process timed out after {timeout}ms";
                    return result;
                }

                result.ExitCode = process.ExitCode;
            }
            catch (Exception ex)
            {
                result.Error = ex.Message;
            }

            return result;
        }

        /// <summary>
        /// Fires and forgets a process (e.g., Windows Update in background).
        /// Returns immediately after starting.
        /// </summary>
        public static bool FireAndForget(string fileName, string arguments)
        {
            try
            {
                var psi = new ProcessStartInfo(fileName, arguments)
                {
                    UseShellExecute = true,
                    CreateNoWindow = false,
                    WindowStyle = ProcessWindowStyle.Normal
                };
                Process.Start(psi);
                return true;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Runs a PowerShell command silently using Base64 encoding to avoid escaping issues.
        /// </summary>
        public static ProcessResult RunPowerShell(
            string script,
            int? timeoutMs = null,
            bool visible = false,
            CancellationToken ct = default)
        {
            var plainTextBytes = Encoding.Unicode.GetBytes(script);
            string base64 = Convert.ToBase64String(plainTextBytes);
            string args = $"-NoProfile -ExecutionPolicy Bypass -EncodedCommand {base64}";

            return visible
                ? RunVisible("powershell", args, timeoutMs, ct)
                : RunSilent("powershell", args, timeoutMs, ct: ct);
        }

        /// <summary>
        /// Runs a simple PowerShell one-liner command silently.
        /// </summary>
        public static ProcessResult RunPowerShellCommand(
            string command,
            int? timeoutMs = null,
            CancellationToken ct = default)
        {
            return RunSilent("powershell", $"-NoProfile -Command \"{command}\"", timeoutMs, ct: ct);
        }
    }
}
