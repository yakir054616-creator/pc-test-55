using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace WinAutomator
{
    /// <summary>
    /// Centralized configuration loaded from appsettings.json.
    /// Falls back to sensible defaults if the file is missing or malformed.
    /// </summary>
    public sealed class AppConfig
    {
        private static AppConfig? _instance;
        private static readonly object _lock = new();

        public BatteryApiConfig BatteryApi { get; set; } = new();
        public TimeoutConfig Timeouts { get; set; } = new();
        public OemConfig Oem { get; set; } = new();
        public CleanupConfig Cleanup { get; set; } = new();
        public string LogFolderName { get; set; } = "WinAutomator_Logs";
        public string[] Manufacturers { get; set; } = { "Lenovo", "Dell", "HP" };

        /// <summary>
        /// Thread-safe singleton accessor. Loads config once from disk.
        /// </summary>
        public static AppConfig Current
        {
            get
            {
                if (_instance == null)
                {
                    lock (_lock)
                    {
                        _instance ??= Load();
                    }
                }
                return _instance;
            }
        }

        /// <summary>
        /// Returns the full path to the log folder on the user's Desktop.
        /// </summary>
        public string GetLogFolderPath()
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string path = Path.Combine(desktop, LogFolderName);
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);
            return path;
        }

        private static AppConfig Load()
        {
            try
            {
                // Look beside the EXE first, then in working directory
                string? exeDir = Path.GetDirectoryName(Environment.ProcessPath);
                string[] searchPaths = new[]
                {
                    Path.Combine(exeDir ?? ".", "appsettings.json"),
                    Path.Combine(AppContext.BaseDirectory, "appsettings.json"),
                    "appsettings.json"
                };

                foreach (string path in searchPaths)
                {
                    if (File.Exists(path))
                    {
                        string json = File.ReadAllText(path);
                        var options = new JsonSerializerOptions
                        {
                            PropertyNameCaseInsensitive = true,
                            ReadCommentHandling = JsonCommentHandling.Skip
                        };
                        return JsonSerializer.Deserialize<AppConfig>(json, options) ?? new AppConfig();
                    }
                }
            }
            catch { /* Fall through to defaults */ }

            return new AppConfig(); // All defaults
        }
    }

    public class BatteryApiConfig
    {
        public string Url { get; set; } = "https://battery-health-checker-d9eac7e3.base44.app/api/apps/6978e65367501f48d9eac7e3/entities/BatteryTest";
    }

    public class TimeoutConfig
    {
        public int DismTimeoutMs { get; set; } = 180000;
        public int WingetTimeoutMs { get; set; } = 600000;
        public int ProcessDefaultTimeoutMs { get; set; } = 300000;
        public int HpiaExtractionTimeoutMs { get; set; } = 120000;
        public int OemPostInstallDelayMs { get; set; } = 2000;
        public int WindowsUpdateWaitMs { get; set; } = 60000;
        public int PhaseTransitionDelayMs { get; set; } = 2000;
    }

    public class OemConfig
    {
        public string[] DellPaths { get; set; } = {
            @"C:\Program Files\Dell\CommandUpdate\dcu-cli.exe",
            @"C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe"
        };
        public string DellWingetId { get; set; } = "Dell.CommandUpdate.Universal";

        public string[] LenovoPaths { get; set; } = {
            @"C:\Program Files (x86)\Lenovo\System Update\tvsu.exe",
            @"C:\Program Files\Lenovo\System Update\tvsu.exe",
            @"C:\Program Files (x86)\Lenovo\System Update\tvsukernel.exe",
            @"C:\Program Files\Lenovo\System Update\tvsukernel.exe"
        };
        public string LenovoWingetId { get; set; } = "Lenovo.SystemUpdate";

        public string[] HpPaths { get; set; } = {
            @"C:\HP_Image_Assistant\HPImageAssistant.exe",
            @"C:\Program Files\HP\HP Image Assistant\HPImageAssistant.exe",
            @"C:\Program Files (x86)\HP\HP Image Assistant\HPImageAssistant.exe"
        };
        public string HpiaDownloadUrl { get; set; } = "https://hpia.hpcloud.hp.com/downloads/hpia/hp-hpia-5.1.0.3.exe";
        public string HpiaExtractFolder { get; set; } = @"C:\HP_Image_Assistant";
        
        public string[] UniversalDriverPaths { get; set; } = {
            @"D:\SDI\sdi.exe",
            @"D:\SDIO\SDIO_auto.bat",
            @"E:\SDI\sdi.exe",
            @"C:\TechTools\SDIO\SDIO_x64_R760.exe"
        };
        public string SdioDownloadUrl { get; set; } = "https://www.glenn.delahoy.com/downloads/sdio/SDIO_1.17.8.829.zip";
    }

    public class CleanupConfig
    {
        public string[] TempPaths { get; set; } = {
            "%TEMP%",
            @"C:\Windows\Temp",
            @"C:\Windows\Prefetch",
            @"C:\Windows\SoftwareDistribution\Download"
        };

        /// <summary>
        /// Resolves environment variables in paths (e.g., %TEMP%).
        /// </summary>
        public string[] GetResolvedPaths()
        {
            var resolved = new string[TempPaths.Length];
            for (int i = 0; i < TempPaths.Length; i++)
                resolved[i] = Environment.ExpandEnvironmentVariables(TempPaths[i]);
            return resolved;
        }
    }
}
