using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Management;
using System.Reflection;
using System.Media;

namespace WinAutomator
{
    public static class AutomationLogic
    {
        // Reusable HttpClient to prevent Socket Exhaustion (TCP port exhaustion)
        private static readonly HttpClient SharedClient = new HttpClient();

        // ----------------- PHASE 1 (Disk + DISM) -----------------

        public static void ExtendCDrive()
        {
            string ps = "$s = (Get-PartitionSupportedSize -DriveLetter C).SizeMax; Resize-Partition -DriveLetter C -Size $s";
            ProcessRunner.RunPowerShellCommand(ps);
        }

        public static bool RunDismCommand(bool isFullAuto)
        {
            var config = AppConfig.Current;
            var result = ProcessRunner.RunVisible("DISM", "/Online /Cleanup-Image /RestoreHealth",
                isFullAuto ? config.Timeouts.DismTimeoutMs : (int?)null);

            return isFullAuto || result.Success;
        }

        // ----------------- PHASE 2 (OS + Updates) -----------------

        public static void ChangeHostname(string newName)
        {
            ProcessRunner.RunPowerShellCommand($"Rename-Computer -NewName '{newName}' -Force -ErrorAction SilentlyContinue");
        }

        public static void ActivateWindows()
        {
            ProcessRunner.RunSilent("cscript", "//b c:\\windows\\system32\\slmgr.vbs /ato");
        }

        public static void ActivateOffice()
        {
            string office64 = @"C:\Program Files\Microsoft Office\Office16\ospp.vbs";
            string office32 = @"C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs";

            if (File.Exists(office64)) ProcessRunner.RunSilent("cscript", $"//b \"{office64}\" /act");
            else if (File.Exists(office32)) ProcessRunner.RunSilent("cscript", $"//b \"{office32}\" /act");
        }

        public static void RunOemUpdates(Action<string> updateStatus, string selectedManufacturer)
        {
            var config = AppConfig.Current;
            string manLower = selectedManufacturer.ToLower();
            const int maxPasses = 2;

            // Determine if this is a GUI-only tool (Lenovo kernel) – only one pass for those
            bool isGuiOnly = manLower.Contains("lenovo") &&
                             !File.Exists(@"C:\Program Files (x86)\Lenovo\System Update\tvsu.exe") &&
                             !File.Exists(@"C:\Program Files\Lenovo\System Update\tvsu.exe");

            int passes = isGuiOnly ? 1 : maxPasses;

            for (int pass = 1; pass <= passes; pass++)
            {
                if (passes > 1)
                    updateStatus($"סיבוב {pass}/{passes} של עדכוני {selectedManufacturer}...");
                else
                    updateStatus($"התחלת עדכוני יצרן עבור: {selectedManufacturer}...");

                try
                {
                    bool hadActivity = false;

                    if (manLower.Contains("dell"))
                        hadActivity = RunDellUpdates(updateStatus);
                    else if (manLower.Contains("lenovo"))
                        hadActivity = RunLenovoUpdates(updateStatus);
                    else if (manLower.Contains("hp") || manLower.Contains("hewlett"))
                        hadActivity = RunHpUpdates(updateStatus);
                    else
                    {
                        updateStatus($"נבחר יצרן '{selectedManufacturer}'. מדלג על כלי יצרן ומשתמש בעדכונים אוניברסליים.");
                        break;
                    }

                    // If first pass had no activity, no need for a second pass
                    if (pass == 1 && !hadActivity && passes > 1)
                    {
                        updateStatus("לא נמצאו עדכונים בסיבוב 1, דילוג על סיבוב 2.");
                        break;
                    }

                    // Between passes: wait 30s for any background operations to settle
                    if (pass < passes)
                    {
                        updateStatus("ממתין 30 שניות לווידוא סיים פעולות רקע לפני סיבוב נוסף...");
                        System.Threading.Thread.Sleep(30000);
                    }
                }
                catch (Exception ex)
                {
                    updateStatus($"שגיאה בעדכוני יצרן (סיבוב {pass}): {ex.Message}");
                }
            }

            updateStatus($"עדכוני יצרן ({selectedManufacturer}) הושלמו ✓");
            System.Threading.Thread.Sleep(config.Timeouts.OemPostInstallDelayMs);
        }


        private static void RunWingetInstall(string packageId, Action<string> updateStatus)
        {
            var config = AppConfig.Current;
            updateStatus($"מתקין {packageId} דרך Winget (עלול לקחת כמה דקות)...");
            string cmd = $"winget install --id {packageId} --silent --accept-package-agreements --accept-source-agreements --force";
            
            var result = ProcessRunner.RunVisible("powershell",
                $"-NoProfile -NonInteractive -Command \"{cmd}\"",
                config.Timeouts.WingetTimeoutMs);

            if (!result.Success && !result.TimedOut)
            {
                // Fallback: try via cmd
                ProcessRunner.RunVisible("cmd", $"/c {cmd}", config.Timeouts.WingetTimeoutMs);
            }
        }

        private static bool RunDellUpdates(Action<string> updateStatus)
        {
            var config = AppConfig.Current;
            updateStatus("מחפש Dell Command Update...");

            string dcuPath = Array.Find(config.Oem.DellPaths, File.Exists) ?? "";
            bool freshlyInstalled = false;
            if (string.IsNullOrEmpty(dcuPath))
            {
                RunWingetInstall(config.Oem.DellWingetId, updateStatus);
                dcuPath = Array.Find(config.Oem.DellPaths, File.Exists) ?? "";
                freshlyInstalled = true;
            }

            if (string.IsNullOrEmpty(dcuPath))
            {
                updateStatus("⚠ Dell Command Update לא נמצא. ממשיך ב-Windows Update.");
                return false;
            }

            // If just installed via winget, wait for initialization to complete
            if (freshlyInstalled)
            {
                updateStatus("ממתין 15 שניות לאתחול Dell Command Update לאחר התקנה...");
                System.Threading.Thread.Sleep(15000);
            }

            string logFolder = config.GetLogFolderPath();
            string logFile = Path.Combine(logFolder, $"DellUpdate_{DateTime.Now:HHmm}.log");

            updateStatus("מריץ Dell Command Update – מתקין עדכונים...");
            // Use cmd /c start /wait to ensure dcu-cli.exe blocks until all child processes complete
            string cmdArgs = $"/c start /wait \"\" \"{dcuPath}\" /applyUpdates -reboot=disable -outputLog=\"{logFile}\"";
            var result = ProcessRunner.RunSilent("cmd", cmdArgs, config.Timeouts.WingetTimeoutMs);

            if (result.TimedOut)
                updateStatus("⚠ Dell Command Update חרג מזמן ההמתנה – בדוק את הלוג");
            else
                updateStatus($"עדכוני Dell הושלמו (Exit: {result.ExitCode}) ✓ נא לבדוק את הלוג");

            // ExitCode 0 = no updates, 1 = updates applied, 5 = reboot needed — all mean "tool ran"
            return !result.TimedOut;
        }

        private static bool RunLenovoUpdates(Action<string> updateStatus)
        {
            var config = AppConfig.Current;
            updateStatus("מחפש Lenovo System Update...");

            string tvsuPath = Array.Find(config.Oem.LenovoPaths, File.Exists) ?? "";
            bool freshlyInstalled = false;
            if (string.IsNullOrEmpty(tvsuPath))
            {
                RunWingetInstall(config.Oem.LenovoWingetId, updateStatus);
                tvsuPath = Array.Find(config.Oem.LenovoPaths, File.Exists) ?? "";
                freshlyInstalled = true;
            }

            if (string.IsNullOrEmpty(tvsuPath))
            {
                updateStatus("⚠ Lenovo System Update לא נמצא. ממשיך ב-Windows Update.");
                return false;
            }

            // If just installed via winget, wait for initialization to complete
            if (freshlyInstalled)
            {
                updateStatus("ממתין 15 שניות לאתחול Lenovo System Update לאחר התקנה...");
                System.Threading.Thread.Sleep(15000);
            }

            string logFolder = config.GetLogFolderPath();
            string logFile = Path.Combine(logFolder, $"LenovoUpdate_{DateTime.Now:HHmm}.log");

            if (tvsuPath.EndsWith("tvsu.exe", StringComparison.OrdinalIgnoreCase))
            {
                updateStatus("מריץ Lenovo System Update (CLI) – מתקין עדכונים...");

                // tvsu.exe spawns child processes and returns immediately.
                // We use cmd /c start /wait to block until child processes finish,
                // and add -nolicense for fully unattended mode.
                string cmdArgs = $"/c start /wait \"\" \"{tvsuPath}\" /CM -search A -action INSTALL -packagetypes 1,2,3 -includerebootpackages 3,4 -noreboot -noicon -nolicense -log \"{logFile}\"";
                var result = ProcessRunner.RunSilent("cmd", cmdArgs, config.Timeouts.WingetTimeoutMs);

                // Also wait for any Lenovo background child processes to complete
                updateStatus("ממתין לסיום תהליכי רקע של Lenovo...");
                WaitForLenovoChildProcesses(60);

                if (result.TimedOut)
                    updateStatus("⚠ Lenovo System Update חרג מזמן ההמתנה – בדוק את הלוג");
                else
                    updateStatus($"עדכוני Lenovo הושלמו (Exit: {result.ExitCode}) ✓");

                return !result.TimedOut;
            }
            else
            {
                // GUI version – open once, technician approves manually
                updateStatus("פותח Lenovo System Update (GUI) – אנא אשר את ההתקנה בחלון שנפתח...");
                ProcessRunner.RunVisible(tvsuPath, "", config.Timeouts.WingetTimeoutMs);
                updateStatus("עדכוני Lenovo הושלמו ✓");
                return false; // GUI mode – no second pass needed
            }
        }

        /// <summary>
        /// Waits for Lenovo System Update child processes (tvsukernel, tvsu, MapDrv) to finish.
        /// tvsu.exe spawns background processes that do the actual work.
        /// </summary>
        private static void WaitForLenovoChildProcesses(int maxWaitSeconds)
        {
            string[] processNames = { "tvsukernel", "tvsu", "MapDrv", "Lenovo.LSU" };
            var sw = System.Diagnostics.Stopwatch.StartNew();

            // Allow 5 seconds for processes to spawn
            System.Threading.Thread.Sleep(5000);

            while (sw.Elapsed.TotalSeconds < maxWaitSeconds)
            {
                bool anyRunning = false;
                foreach (string name in processNames)
                {
                    try
                    {
                        var procs = System.Diagnostics.Process.GetProcessesByName(name);
                        if (procs.Length > 0)
                        {
                            anyRunning = true;
                            foreach (var p in procs) p.Dispose();
                            break;
                        }
                        foreach (var p in procs) p.Dispose();
                    }
                    catch { }
                }

                if (!anyRunning) break;
                System.Threading.Thread.Sleep(3000);
            }
        }

        private static bool RunHpUpdates(Action<string> updateStatus)
        {
            var config = AppConfig.Current;
            updateStatus("מחפש HP Image Assistant (HPIA)...");

            string logFolder = config.GetLogFolderPath();
            string hpPath = Array.Find(config.Oem.HpPaths, File.Exists) ?? "";
            bool freshlyInstalled = false;

            // Step 1: try winget
            if (string.IsNullOrEmpty(hpPath))
            {
                updateStatus("HPIA לא נמצא. מנסה התקנה דרך Winget...");
                RunWingetInstall("HP.HPImageAssistant", updateStatus);
                hpPath = Array.Find(config.Oem.HpPaths, File.Exists) ?? "";
                freshlyInstalled = true;
            }

            // Step 2: direct download of the portable exe (no extraction needed)
            if (string.IsNullOrEmpty(hpPath))
            {
                updateStatus("מוריד HP Image Assistant ישירות מ-HP...");
                string portableExe = Path.Combine(config.Oem.HpiaExtractFolder, "HPImageAssistant.exe");
                try
                {
                    if (!Directory.Exists(config.Oem.HpiaExtractFolder))
                        Directory.CreateDirectory(config.Oem.HpiaExtractFolder);

                    using (var client = new System.Net.WebClient())
                        client.DownloadFile("https://ftp.hp.com/pub/caps-softpaq/cmit/HPImageAssistant.exe", portableExe);

                    if (File.Exists(portableExe))
                    {
                        hpPath = portableExe;
                        freshlyInstalled = true;
                    }
                    else
                        updateStatus("⚠ ההורדה הסתיימה אך הקובץ לא נמצא.");
                }
                catch (Exception ex) { updateStatus($"⚠ הורדת HPIA נכשלה: {ex.Message}"); }
            }

            // Step 3: run HPIA if found
            if (string.IsNullOrEmpty(hpPath))
            {
                updateStatus("⚠ HP Image Assistant לא זמין. ממשיך ב-Windows Update.");
                return false;
            }

            // If just installed, wait for initialization
            if (freshlyInstalled)
            {
                updateStatus("ממתין 10 שניות לאתחול HP Image Assistant...");
                System.Threading.Thread.Sleep(10000);
            }

            updateStatus("מריץ HP Image Assistant – מנתח ומתקין עדכונים...");
            // Use cmd /c start /wait to ensure HPIA blocks until all child processes complete
            string cmdArgs = $"/c start /wait \"\" \"{hpPath}\" /Operation:Analyze /Action:Install /Silent /Category:All /ReportFolder:\"{logFolder}\"";
            var result = ProcessRunner.RunSilent("cmd", cmdArgs, config.Timeouts.WingetTimeoutMs);

            if (result.TimedOut)
                updateStatus("⚠ HPIA חרג מזמן ההמתנה – בדוק את הלוג");
            else
                updateStatus($"עדכוני HP הושלמו (Exit: {result.ExitCode}) ✓ נא לבדוק את הלוג");

            return !result.TimedOut;
        }

        public static bool CheckForMissingDrivers()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PnPEntity WHERE ConfigManagerErrorCode <> 0"))
                {
                    return searcher.Get().Count > 0;
                }
            }
            catch
            {
                return false; 
            }
        }

        public static void RunUniversalDriverUpdate(Action<string> updateStatus)
        {
            var config = AppConfig.Current;
            string sdioPath = Array.Find(config.Oem.UniversalDriverPaths, File.Exists) ?? "";

            // If SDIO not found locally, download portable version from the internet
            if (string.IsNullOrEmpty(sdioPath))
            {
                updateStatus("SDIO לא נמצא מקומית – מוריד גרסה ניידת מהאינטרנט...");
                string sdioFolder = @"C:\TechTools\SDIO";
                string downloadedZip = Path.Combine(sdioFolder, "SDIO_latest.zip");
                
                try
                {
                    if (!Directory.Exists(sdioFolder))
                        Directory.CreateDirectory(sdioFolder);

                    updateStatus("מוריד SDIO מהאתר הרשמי...");
                    using (var client = new System.Net.WebClient())
                        client.DownloadFile(config.Oem.SdioDownloadUrl, downloadedZip);
                    
                    if (File.Exists(downloadedZip))
                    {
                        updateStatus("חולץ קובץ SDIO...");
                        System.IO.Compression.ZipFile.ExtractToDirectory(downloadedZip, sdioFolder, true);
                        try { File.Delete(downloadedZip); } catch { }

                        // Search for the SDIO executable inside extracted folder
                        string[] candidates = Directory.GetFiles(sdioFolder, "SDIO_x64*.exe", SearchOption.AllDirectories);
                        if (candidates.Length == 0)
                            candidates = Directory.GetFiles(sdioFolder, "SDIO*.exe", SearchOption.AllDirectories);

                        if (candidates.Length > 0)
                        {
                            sdioPath = candidates[0];
                            updateStatus($"SDIO הורד וחולץ בהצלחה ✓ ({Path.GetFileName(sdioPath)})");
                        }
                        else
                        {
                            updateStatus("⚠ הארכיון חולץ אך לא נמצא קובץ SDIO_x64*.exe.");
                        }
                    }
                    else
                    {
                        updateStatus("⚠ ההורדה הסתיימה אך הקובץ לא נמצא.");
                    }
                }
                catch (Exception ex)
                {
                    updateStatus($"⚠ הורדת SDIO נכשלה: {ex.Message}. ממשיך ללא מנוע אוניברסלי.");
                    return;
                }
            }

            if (string.IsNullOrEmpty(sdioPath))
            {
                updateStatus("⚠ לא נמצא כלי התקנת דרייברים אוניברסלי (SDIO). ממשיך...");
                return;
            }

            updateStatus($"מפעיל מנוע דרייברים אוניברסלי ({Path.GetFileName(sdioPath)})... מחפש עדכונים באינטרנט ומתקין.");
            
            // -checkupdates: download latest online driver index
            // -autoinstall: automatically install missing/outdated drivers
            // -autoclose: close when done
            var result = ProcessRunner.RunVisible(sdioPath, "-checkupdates -autoinstall -autoclose", config.Timeouts.WingetTimeoutMs);

            if (result.TimedOut)
                updateStatus("⚠ מנוע העדכונים האוניברסלי חרג מזמן ההמתנה");
            else
                updateStatus($"השלמת דרייברים אוניברסלית הושלמה (כולל מאגר אונליין) ✓");
        }

        public static void RunWindowsUpdates(Action<string> updateStatus)
        {
            var config = AppConfig.Current;
            try
            {
                updateStatus("מגדיר סביבת עדכוני Windows (דו\"ח נשמר בשולחן מעבודה)...");
                string setupPs = "Set-ExecutionPolicy Bypass -Scope Process -Force; " +
                    "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; " +
                    "Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -ErrorAction SilentlyContinue; " +
                    "Install-Module PSWindowsUpdate -Force -Confirm:$false -ErrorAction SilentlyContinue";
                ProcessRunner.RunSilent("powershell", $"-NoProfile -WindowStyle Hidden -Command \"{setupPs}\"");

                updateStatus("מוריד ומתקין עדכוני Windows. נפתח חלון כחול להצגת ההתקדמות...");

                string logFolder = config.GetLogFolderPath();
                string logFile = Path.Combine(logFolder, $"WindowsUpdate_{DateTime.Now:HHmm}.log");

                string updatePs = $"Get-WindowsUpdate -AcceptAll -Install -AutoReboot:$false -IgnoreReboot -ErrorAction SilentlyContinue | Out-File -FilePath '{logFile}' -Encoding utf8 -Append";

                ProcessRunner.FireAndForget("powershell", $"-NoProfile -Command \"{updatePs}\"");

                updateStatus("עדכוני Windows הופעלו במקביל ✓ החלון הכחול ממשיך ברקע, אנחנו ממשיכים הלאה!");
            }
            catch (Exception ex)
            {
                updateStatus($"אזהרה: הפעלת עדכוני Windows נכשלה: {ex.Message}. ממשיכים...");
            }
        }

        // ----------------- PHASE 3 (Diagnostics & API) -----------------

        public static async Task<double> PerformBatteryReportAndApi(string techName, string serialNum, string cpuGen, Action<string> log)
        {
            var config = AppConfig.Current;
            log("מכין דו\"ח סוללה...");
            try
            {
                string tempXml = Path.Combine(Path.GetTempPath(), "battery.xml");
                ProcessRunner.RunSilent("powercfg", $"/batteryreport /xml /output \"{tempXml}\"");

                if (!File.Exists(tempXml)) { log("שגיאה בהפקת דו\"ח XML."); return -1; }

                string xmlContent = File.ReadAllText(tempXml);
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(xmlContent);

                double original = 0;
                double current = 0;
                int cycles = 0;

                XmlNode batteryNode = doc.GetElementsByTagName("Battery")[0];
                if (batteryNode != null)
                {
                    foreach (XmlNode child in batteryNode.ChildNodes)
                    {
                        if (child.Name == "DesignCapacity") double.TryParse(child.InnerText, out original);
                        if (child.Name == "FullChargeCapacity") double.TryParse(child.InnerText, out current);
                        if (child.Name == "CycleCount") int.TryParse(child.InnerText, out cycles);
                    }
                }

                if (original <= 0) original = 1;
                double health = (current / original) * 100.0;
                if (health > 100) health = 100;
                if (health < 0) health = 0;

                log("זורק נתונים ל-API כעת...");

                var payload = new
                {
                    original_capacity = original,
                    current_capacity = current,
                    health_percentage = Math.Round(health, 2),
                    test_date = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    technician_name = techName,
                    serial_number = serialNum,
                    cpu_generation = cpuGen,
                    cycles = cycles
                };

                var json = JsonSerializer.Serialize(payload);
                var content = new StringContent(json, Encoding.UTF8, "application/json");
                await SharedClient.PostAsync(config.BatteryApi.Url, content);

                log("הדו\"ח שוגר למסד בהצלחה!");
                try { File.Delete(tempXml); } catch { }
                return health;
            }
            catch (Exception ex)
            {
                log("שגיאה במנגנון API הסוללה: " + ex.Message);
                return -1;
            }
        }

        public static string GetSsdHealth()
        {
            try
            {
                var scope = new ManagementScope(@"\\localhost\ROOT\Microsoft\Windows\Storage");
                var query = new ObjectQuery("SELECT * FROM MSFT_PhysicalDisk");
                var searcher = new ManagementObjectSearcher(scope, query);

                foreach (ManagementObject disk in searcher.Get())
                {
                    string health = disk["HealthStatus"]?.ToString() == "0" ? "Healthy (תקין)" : "Warning (אזהרה)";

                    var relSearcher = new ManagementObjectSearcher(scope, new ObjectQuery($"SELECT * FROM MSFT_StorageReliabilityCounter WHERE ObjectId LIKE '%{disk["DeviceId"]}%'"));
                    foreach (ManagementObject rel in relSearcher.Get())
                    {
                        if (rel["Wear"] != null)
                        {
                            return $"{health} | Wear: {rel["Wear"]}% Left";
                        }
                    }
                    return health;
                }
            }
            catch (Exception ex)
            {
                return $"Status: Unknown (WMI Error: {ex.Message})";
            }
            return "Status: Unknown";
        }

        // ----------------- DEFRAG / OPTIMIZE SSD -----------------

        /// <summary>
        /// Runs defrag /O on drive C: which auto-detects disk type:
        /// SSD -> TRIM optimization, HDD -> traditional defragmentation.
        /// Returns the output text for logging.
        /// </summary>
        public static string OptimizeDrive(Action<string> log)
        {
            log("מבצע אופטימיזציה לכונן C (TRIM ל-SSD / דיפרוג ל-HDD)...");
            var result = ProcessRunner.RunSilent("defrag", "C: /O /U", redirectOutput: true);

            if (result.TimedOut)
            {
                log("אופטימיזציית כונן חרגה מזמן המתנה");
                return "Timeout";
            }

            if (result.Success)
                log("אופטימיזציית כונן הושלמה בהצלחה ✓");
            else
                log($"אופטימיזציית כונן הסתיימה עם אזהרה (Exit: {result.ExitCode})");

            return result.Output;
        }

        // ----------------- VICTORY SOUND -----------------

        /// <summary>
        /// Generates and plays a short victory fanfare sound (3 ascending tones).
        /// No external files needed – pure in-memory WAV synthesis.
        /// </summary>
        public static void PlayVictorySound()
        {
            try
            {
                // Three ascending notes: C5 (523Hz), E5 (659Hz), G5 (784Hz) – a major chord arpeggio
                int[][] notes = new int[][] {
                    new int[] { 523, 200 },  // C5 - 200ms
                    new int[] { 659, 200 },  // E5 - 200ms
                    new int[] { 784, 400 },  // G5 - 400ms (longer final note)
                };

                int sampleRate = 44100;
                int totalSamples = 0;
                foreach (var note in notes)
                    totalSamples += (int)(sampleRate * note[1] / 1000.0);

                int byteCount = totalSamples * 4; // 16-bit stereo
                byte[] buffer = new byte[44 + byteCount];

                // RIFF WAV header
                Array.Copy(Encoding.ASCII.GetBytes("RIFF"), 0, buffer, 0, 4);
                BitConverter.GetBytes(36 + byteCount).CopyTo(buffer, 4);
                Array.Copy(Encoding.ASCII.GetBytes("WAVE"), 0, buffer, 8, 4);
                Array.Copy(Encoding.ASCII.GetBytes("fmt "), 0, buffer, 12, 4);
                BitConverter.GetBytes(16).CopyTo(buffer, 16);
                BitConverter.GetBytes((short)1).CopyTo(buffer, 20);   // PCM
                BitConverter.GetBytes((short)2).CopyTo(buffer, 22);   // Stereo
                BitConverter.GetBytes(sampleRate).CopyTo(buffer, 24);
                BitConverter.GetBytes(sampleRate * 4).CopyTo(buffer, 28);
                BitConverter.GetBytes((short)4).CopyTo(buffer, 32);
                BitConverter.GetBytes((short)16).CopyTo(buffer, 34);
                Array.Copy(Encoding.ASCII.GetBytes("data"), 0, buffer, 36, 4);
                BitConverter.GetBytes(byteCount).CopyTo(buffer, 40);

                int offset = 44;
                foreach (var note in notes)
                {
                    int freq = note[0];
                    int durationMs = note[1];
                    int samples = (int)(sampleRate * durationMs / 1000.0);

                    for (int i = 0; i < samples; i++)
                    {
                        double envelope = 1.0;
                        int attackSamples = sampleRate / 100;
                        int releaseSamples = sampleRate / 50;
                        if (i < attackSamples) envelope = (double)i / attackSamples;
                        if (i > samples - releaseSamples) envelope = (double)(samples - i) / releaseSamples;

                        double sample = Math.Sin(2 * Math.PI * freq * i / sampleRate) * 12000 * envelope;
                        short s = (short)sample;
                        BitConverter.GetBytes(s).CopyTo(buffer, offset);
                        BitConverter.GetBytes(s).CopyTo(buffer, offset + 2);
                        offset += 4;
                    }
                }

                using (var ms = new MemoryStream(buffer))
                {
                    using (var sp = new System.Media.SoundPlayer(ms))
                    {
                        sp.PlaySync();
                    }
                }
            }
            catch { } // Fail silently – sound is a bonus, not critical
        }
        public static void PlaySongTest()
        {
            try
            {
                string tempPath = Path.Combine(Path.GetTempPath(), "WinAutomator_SpeakerTest.mp3");
                
                // Extract embedded resource to temp file if it doesn't exist or to ensure it's fresh
                var assembly = Assembly.GetExecutingAssembly();
                string resourceName = "WinAutomator.SpeakerTest.mp3";
                
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        using (FileStream fileStream = new FileStream(tempPath, FileMode.Create, FileAccess.Write))
                        {
                            stream.CopyTo(fileStream);
                        }
                    }
                }

                if (!File.Exists(tempPath)) return;

                mciSendString("close testSong", null, 0, IntPtr.Zero);
                mciSendString($"open \"{tempPath}\" alias testSong", null, 0, IntPtr.Zero);
                
                // Set application volume to 80% (800 / 1000)
                mciSendString("setaudio testSong volume to 800", null, 0, IntPtr.Zero);
                
                // Play the song
                mciSendString("play testSong", null, 0, IntPtr.Zero);
            }
            catch { }
        }

        public static void SetSongPanning(int leftVol, int rightVol)
        {
            try
            {
                mciSendString($"setaudio testSong left volume to {leftVol}", null, 0, IntPtr.Zero);
                mciSendString($"setaudio testSong right volume to {rightVol}", null, 0, IntPtr.Zero);
            }
            catch { }
        }

        public static void StopSongTest()
        {
            try
            {
                mciSendString("stop testSong", null, 0, IntPtr.Zero);
                mciSendString("close testSong", null, 0, IntPtr.Zero);
            }
            catch { }
        }

        public static void PlayStereoTest(bool leftSide)
        {
            byte[] wav = CreateBeepWav(440, 3.0, leftSide);
            using (var ms = new MemoryStream(wav))
            {
                using (var sp = new System.Media.SoundPlayer(ms))
                {
                    sp.PlaySync();
                }
            }
        }

        private static byte[] CreateBeepWav(int freq, double duration, bool leftSide)
        {
            int sampleRate = 44100;
            int samples = (int)(sampleRate * duration);
            int byteCount = samples * 4; // 16-bit stereo
            byte[] buffer = new byte[44 + byteCount];

            // RIFF Header
            Array.Copy(Encoding.ASCII.GetBytes("RIFF"), 0, buffer, 0, 4);
            BitConverter.GetBytes(36 + byteCount).CopyTo(buffer, 4);
            Array.Copy(Encoding.ASCII.GetBytes("WAVE"), 0, buffer, 8, 4);
            Array.Copy(Encoding.ASCII.GetBytes("fmt "), 0, buffer, 12, 4);
            BitConverter.GetBytes(16).CopyTo(buffer, 16);
            BitConverter.GetBytes((short)1).CopyTo(buffer, 20); // PCM
            BitConverter.GetBytes((short)2).CopyTo(buffer, 22); // Stereo
            BitConverter.GetBytes(sampleRate).CopyTo(buffer, 24);
            BitConverter.GetBytes(sampleRate * 4).CopyTo(buffer, 28);
            BitConverter.GetBytes((short)4).CopyTo(buffer, 32);
            BitConverter.GetBytes((short)16).CopyTo(buffer, 34);
            Array.Copy(Encoding.ASCII.GetBytes("data"), 0, buffer, 36, 4);
            BitConverter.GetBytes(byteCount).CopyTo(buffer, 40);

            for (int i = 0; i < samples; i++)
            {
                short s = (short)(Math.Sin(2 * Math.PI * freq * i / sampleRate) * 32767);
                if (leftSide)
                {
                    BitConverter.GetBytes(s).CopyTo(buffer, 44 + i * 4); // Left
                    BitConverter.GetBytes((short)0).CopyTo(buffer, 44 + i * 4 + 2); // Right
                }
                else
                {
                    BitConverter.GetBytes((short)0).CopyTo(buffer, 44 + i * 4); // Left
                    BitConverter.GetBytes(s).CopyTo(buffer, 44 + i * 4 + 2); // Right
                }
            }
            return buffer;
        }

        public static void PerformSystemCleanup(Action<string> log)
        {
            log("מנקה קבצים זמניים ומטמון עדכונים...");
            string[] paths = AppConfig.Current.Cleanup.GetResolvedPaths();

            foreach (string path in paths)
            {
                try
                {
                    if (!Directory.Exists(path)) continue;
                    DirectoryInfo di = new DirectoryInfo(path);
                    foreach (FileInfo file in di.GetFiles()) { try { file.Delete(); } catch { } }
                    foreach (DirectoryInfo dir in di.GetDirectories()) { try { dir.Delete(true); } catch { } }
                }
                catch { }
            }
        }

        // MIC
        [DllImport("winmm.dll")]
        private static extern long mciSendString(string command, StringBuilder retstring, int ReturnLength, IntPtr callback);

        public static void StartMicRecording()
        {
            mciSendString("open new Type waveaudio Alias recsound", null, 0, IntPtr.Zero);
            mciSendString("record recsound", null, 0, IntPtr.Zero);
        }

        public static void StopAndPlayMic()
        {
            mciSendString("save recsound record.wav", null, 0, IntPtr.Zero);
            mciSendString("close recsound", null, 0, IntPtr.Zero);
            
            // Re-open explicitly to play at 100% volume
            mciSendString("open record.wav alias playMic", null, 0, IntPtr.Zero);
            mciSendString("setaudio playMic volume to 1000", null, 0, IntPtr.Zero); // 100%
            mciSendString("play playMic wait", null, 0, IntPtr.Zero);
            mciSendString("close playMic", null, 0, IntPtr.Zero);
        }

        // CAMERA
        public static void OpenCameraFor30Secs()
        {
            ProcessRunner.FireAndForget("cmd", "/c start microsoft.windows.camera:");
        }

        public static void CloseCamera()
        {
            ProcessRunner.RunSilent("taskkill", "/IM WindowsCamera.exe /F");
        }

        // ----------------- STATE MACHINE & SYSTEM -----------------

        public static void SetupAutoResumeAndRestart(int targetPhase, string tech, string serial, string cpu, bool isFullAuto, bool skipHostname, string manufacturer)
        {
            // Use a persistent folder (AppData) for the state file
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string winAutoDir = Path.Combine(appData, "WinAutomator");
            if (!Directory.Exists(winAutoDir)) Directory.CreateDirectory(winAutoDir);

            string statePath = Path.Combine(winAutoDir, "WinAutomatorState.txt");
            File.WriteAllText(statePath, $"{tech}|{serial}|{isFullAuto}|{targetPhase}|{cpu}|{skipHostname}|{manufacturer}");

            // Best way to find the current EXE path in .NET 8
            string exePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName ?? "";
            if (string.IsNullOrEmpty(exePath)) return;

            // ONLY USE: Scheduled Task (Highest Privilege, Zero Delay, Battery Independent)
            try
            {
                string taskName = "WinAutomator_Resume";
                string psScript = $@"
                    $ErrorActionPreference = 'Stop'
                    Unregister-ScheduledTask -TaskName '{taskName}' -Confirm:$false -ErrorAction SilentlyContinue
                    $action = New-ScheduledTaskAction -Execute '{exePath}' -Argument '--resume'
                    $trigger = New-ScheduledTaskTrigger -AtLogOn
                    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit 0 -Priority 0 -MultipleInstances IgnoreNew
                    Register-ScheduledTask -TaskName '{taskName}' -Action $action -Trigger $trigger -Settings $settings -RunLevel Highest -Force
                ";

                ProcessRunner.RunPowerShell(psScript);
            }
            catch { }

            // Restart the machine
            ProcessRunner.RunSilent("shutdown", "/r /t 0 /f");
        }

        public static void CleanupAutoResume()
        {
            try
            {
                // Delete task
                ProcessRunner.RunSilent("schtasks", "/delete /tn \"WinAutomator_Resume\" /f");

                // Cleanup Startup
                string startupFolder = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
                string batPath = Path.Combine(startupFolder, "WinAutomator_Resume.bat");
                if (File.Exists(batPath)) File.Delete(batPath);

                // RunOnce cleanup
                try {
                    using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", true))
                    {
                        if (key != null && key.GetValue("WinAutomator_Resume") != null)
                        {
                            try { key.DeleteValue("WinAutomator_Resume"); } catch { }
                        }
                    }
                } catch {}

                // Delete the state file
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string statePath = Path.Combine(appData, "WinAutomator", "WinAutomatorState.txt");
                if (File.Exists(statePath)) File.Delete(statePath);
            }
            catch { }
        }

        // ----------------- HARDWARE INFO QUERIES -----------------

        public static string GetCpuGeneration()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("select Name from Win32_Processor"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        string name = obj["Name"]?.ToString() ?? "";
                        if (string.IsNullOrEmpty(name)) continue;

                        string[] parts = { "i3", "i5", "i7", "i9" };
                        foreach (var part in parts)
                        {
                            int idx = name.IndexOf(part, StringComparison.OrdinalIgnoreCase);
                            if (idx != -1)
                            {
                                string sub = name.Substring(idx + part.Length);
                                string digits = "";
                                foreach (char c in sub)
                                {
                                    if (char.IsDigit(c)) digits += c;
                                    else if (digits.Length > 0) break;
                                }

                                if (digits.Length > 0)
                                {
                                    if (digits.Length >= 4 && digits.StartsWith("1")) return digits.Substring(0, 2);
                                    if (digits.Length >= 4) return digits.Substring(0, 1);
                                    if (digits.Length == 3) return digits.Substring(0, 1);
                                    return digits.Substring(0, 1);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
            return "";
        }

        public static string GetCpuName()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("select Name from Win32_Processor"))
                {
                    foreach (var obj in searcher.Get())
                    {
                        string name = obj["Name"]?.ToString() ?? "";
                        if (name.Contains("i9")) return "i9";
                        if (name.Contains("i7")) return "i7";
                        if (name.Contains("i5")) return "i5";
                        if (name.Contains("i3")) return "i3";
                        if (name.Contains("Ryzen 9")) return "Ryzen 9";
                        if (name.Contains("Ryzen 7")) return "Ryzen 7";
                        if (name.Contains("Ryzen 5")) return "Ryzen 5";
                        if (name.Contains("Ryzen 3")) return "Ryzen 3";
                        return name.Split(' ')[0];
                    }
                }
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
            return "";
        }

        public static string GetRamSizeGB()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT TotalPhysicalMemory FROM Win32_ComputerSystem"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        long bytes = Convert.ToInt64(obj["TotalPhysicalMemory"]);
                        double gb = bytes / (1024.0 * 1024.0 * 1024.0);
                        return Math.Round(gb).ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                return $"Error: {ex.Message}";
            }
            return "";
        }

        public static string GetTotalDiskSizeGB()
        {
            try
            {
                using (var searcher = new ManagementObjectSearcher("SELECT Size, MediaType, InterfaceType FROM Win32_DiskDrive WHERE DeviceID LIKE '%PHYSICALDRIVE0%'"))
                {
                    foreach (ManagementObject obj in searcher.Get())
                    {
                        long bytes = Convert.ToInt64(obj["Size"]);
                        double gb = bytes / (1000.0 * 1000.0 * 1000.0); // Disks use metric GB
                        if (gb < 150) return "128";
                        if (gb < 300) return "256";
                        if (gb < 600) return "512";
                        if (gb < 1200) return "1000";
                        return Math.Round(gb).ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                return $"256"; // Silent fallback on error
            }
            return "256"; // Fallback realistic default
        }
    }
}
