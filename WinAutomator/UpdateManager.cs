using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;

namespace WinAutomator
{
    public class UpdateManager
    {
        private static readonly HttpClient _httpClient = new HttpClient();

        public static async Task CheckAndApplyUpdateAsync()
        {
            var config = AppConfig.Current.Updates;
            if (!config.AutoUpdateEnabled) return;

            try
            {
                // 1. Get current version
                var currentVersion = Assembly.GetExecutingAssembly().GetName().Version;
                if (currentVersion == null) return;

                // 2. Query GitHub API
                string url = $"https://api.github.com/repos/{config.GitHubRepoOwner}/{config.GitHubRepoName}/releases";
                _httpClient.DefaultRequestHeaders.Add("User-Agent", "WinAutomator-Updater");
                
                var response = await _httpClient.GetStringAsync(url);
                using var jsonDoc = JsonDocument.Parse(response);
                
                // 3. Find latest release based on channel
                var releases = jsonDoc.RootElement.EnumerateArray();
                JsonElement? targetRelease = null;

                if (config.UpdateChannel.Equals("Beta", StringComparison.OrdinalIgnoreCase))
                {
                    // For Beta, just take the absolute latest release (stable or pre-release)
                    targetRelease = releases.FirstOrDefault();
                }
                else
                {
                    // For Stable, take the latest one that is NOT a pre-release
                    targetRelease = releases.FirstOrDefault(r => !r.GetProperty("prerelease").GetBoolean());
                }

                if (targetRelease == null) return;

                string tagName = targetRelease.Value.GetProperty("tag_name").GetString() ?? "";
                var remoteVersion = ParseVersion(tagName);

                if (remoteVersion > currentVersion)
                {
                    // 4. Find the .exe asset
                    var assets = targetRelease.Value.GetProperty("assets").EnumerateArray();
                    var exeAsset = assets.FirstOrDefault(a => a.GetProperty("name").GetString()?.EndsWith(".exe", StringComparison.OrdinalIgnoreCase) == true);

                    if (exeAsset.ValueKind != JsonValueKind.Undefined)
                    {
                        string downloadUrl = exeAsset.GetProperty("browser_download_url").GetString() ?? "";
                        await PerformUpdate(downloadUrl, tagName);
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Update check failed: {ex.Message}");
            }
        }

        private static Version ParseVersion(string tagName)
        {
            // Clean "v1.0.0" -> "1.0.0"
            string clean = new string(tagName.Where(c => char.IsDigit(c) || c == '.').ToArray());
            return Version.TryParse(clean, out var v) ? v : new Version(0, 0, 0);
        }

        private static async Task PerformUpdate(string downloadUrl, string newVersion)
        {
            try
            {
                string currentExePath = Process.GetCurrentProcess().MainModule?.FileName ?? "";
                if (string.IsNullOrEmpty(currentExePath)) return;

                string tempExePath = Path.Combine(Path.GetDirectoryName(currentExePath) ?? "", "WinAutomator_new.exe");

                // Download the new version
                byte[] data = await _httpClient.GetByteArrayAsync(downloadUrl);
                await File.WriteAllBytesAsync(tempExePath, data);

                // Create the updater batch file
                string batPath = Path.Combine(Path.GetDirectoryName(currentExePath) ?? "", "updater.bat");
                string batContent = $@"
@echo off
timeout /t 2 /nobreak > nul
taskkill /F /IM ""{Path.GetFileName(currentExePath)}"" > nul 2>&1
del /f /q ""{currentExePath}""
move /y ""{tempExePath}"" ""{currentExePath}""
start """" ""{currentExePath}""
del ""%~f0""
";
                await File.WriteAllTextAsync(batPath, batContent);

                // Execute batch and exit
                Process.Start(new ProcessStartInfo
                {
                    FileName = batPath,
                    CreateNoWindow = true,
                    UseShellExecute = false,
                    WindowStyle = ProcessWindowStyle.Hidden
                });

                Environment.Exit(0);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Update application failed: {ex.Message}");
            }
        }
    }
}
