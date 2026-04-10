using System;
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Management;

namespace WinAutomator
{
    public static class AutomationLogic
    {
        // ----------------- PHASE 1 (Disk + DISM) -----------------

        public static void ExtendCDrive()
        {
            string ps = "Get-PartitionSupportedSize -DriveLetter C | Select-Object -ExpandProperty SizeMax | Resize-Partition -DriveLetter C -Size $_";
            RunProcess("powershell", $"-Command \"{ps}\"");
        }

        public static bool RunDismCommand(bool isFullAuto)
        {
            ProcessStartInfo psi = new ProcessStartInfo("DISM", "/Online /Cleanup-Image /RestoreHealth") 
            { CreateNoWindow = false, UseShellExecute = true };

            using (Process p = Process.Start(psi))
            {
                if (isFullAuto)
                {
                    p.WaitForExit(180000); // 3 minutes timeout
                    return true;
                }
                else
                {
                    p.WaitForExit();
                    return p.ExitCode == 0;
                }
            }
        }

        // ----------------- PHASE 2 (OS + Updates) -----------------

        public static void ChangeHostname(string newName)
        {
            RunProcess("wmic", $"computersystem where name=\"%computername%\" call rename name=\"{newName}\"");
        }

        public static void ActivateWindows()
        {
            RunProcess("cscript", "//b c:\\windows\\system32\\slmgr.vbs /ato");
        }

        public static void ActivateOffice()
        {
            string office64 = @"C:\Program Files\Microsoft Office\Office16\ospp.vbs";
            string office32 = @"C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs";

            if (File.Exists(office64)) RunProcess("cscript", $"//b \"{office64}\" /act");
            else if (File.Exists(office32)) RunProcess("cscript", $"//b \"{office32}\" /act");
        }

        public static void RunOemUpdates(Action<string> updateStatus)
        {
            updateStatus("ממתין לסריקת תצורת חומרת המחשב (WMI)...");
            string manufacturer = GetManufacturer().Trim();
            
            try
            {
                string manLower = manufacturer.ToLower();
                if (manLower.Contains("dell"))
                {
                    updateStatus($"זוהה יצרן: {manufacturer}. מחפש עדכוני Dell...");
                    string dcuPath = @"C:\Program Files (x86)\Dell\CommandUpdate\dcu-cli.exe";
                    if (File.Exists(dcuPath))
                        RunProcess(dcuPath, "/applyUpdates -reboot=disable -silent");
                    else
                        updateStatus($"תוכנת העדכונים של מחשבי Dell לא מותקנת - אין עדכוני יצרן.");
                }
                else if (manLower.Contains("lenovo"))
                {
                    updateStatus($"זוהה יצרן: {manufacturer}. מחפש עדכוני Lenovo...");
                    string tvsuPath = @"C:\Program Files (x86)\Lenovo\System Update\tvsu.exe";
                    if (File.Exists(tvsuPath))
                        RunProcess(tvsuPath, "/CM -search A -action INSTALL -packagetypes 1,2,3 -includerebootpackages 1,3,4 -noreboot -noicon -nolicense");
                    else
                        updateStatus($"תוכנת העדכונים של Lenovo TVSU לא מותקנת - אין עדכוני יצרן.");
                }
                else if (manLower.Contains("hp") || manLower.Contains("hewlett"))
                {
                    updateStatus($"זוהה יצרן: {manufacturer}. סורק סביבת HP...");
                    // Basic fallback for HP
                    string hpPath = @"C:\Program Files (x86)\HP\HP Image Assistant\HPImageAssistant.exe";
                    if (File.Exists(hpPath))
                        RunProcess(hpPath, "/Operation:Analyze /Action:Install /Silent");
                    else
                        updateStatus($"תוכנת העדכונים של HP לא מותקנת במחשב - אין עדכוני יצרן.");
                }
                else
                {
                    string toShow = string.IsNullOrEmpty(manufacturer) ? "יצרן לא ידוע" : manufacturer;
                    updateStatus($"זוהה יצרן: {toShow}. לא מצאנו אוטומציה ליצרן זה - אין עדכוני יצרן.");
                }

                // Add an artificial delay so the user can read the output
                System.Threading.Thread.Sleep(3000);
            }
            catch (Exception ex)
            {
                updateStatus($"שגיאה בזיהוי יצרן: {ex.Message}");
                System.Threading.Thread.Sleep(2000);
            }
        }

        public static void RunWindowsUpdates(Action<string> updateStatus)
        {
            updateStatus("מגדיר סביבת עדכוני Windows (NuGet & PSWindowsUpdate)...");
            // Set execution policy and install NuGet provider if needed
            string setupPs = "Set-ExecutionPolicy Bypass -Scope Process -Force; [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force";
            RunProcess("powershell", $"-Command \"{setupPs}\"");

            updateStatus("מתחיל למשוך חבילות אבטחה של Windows Updates. אנא המתן...");
            string updatePs = "Install-Module PSWindowsUpdate -Force -Confirm:$false; Get-WindowsUpdate -AcceptAll -Install -AutoReboot:$false -IgnoreReboot";
            RunProcess("powershell", $"-Command \"{updatePs}\"");
        }

        private static string GetManufacturer()
        {
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("wmic", "computersystem get manufacturer") { RedirectStandardOutput = true, UseShellExecute = false, CreateNoWindow = true };
                using (Process p = Process.Start(psi))
                {
                    string output = p.StandardOutput.ReadToEnd();
                    return output;
                }
            }
            catch { return "Unknown"; }
        }

        // ----------------- PHASE 3 (Diagnostics & API) -----------------

        public static async Task PerformBatteryReportAndApi(string techName, string serialNum, string cpuGen, Action<string> log)
        {
            log("מכין דו\"ח סוללה...");
            try
            {
                string tempXml = Path.Combine(Path.GetTempPath(), "battery.xml");
                RunProcess("powercfg", $"/batteryreport /xml /output \"{tempXml}\"");
                
                if (!File.Exists(tempXml)) { log("שגיאה בהפקת דו\"ח XML."); return; }

                string xmlContent = File.ReadAllText(tempXml);
                XmlDocument doc = new XmlDocument();
                doc.LoadXml(xmlContent);

                double original = 0;
                double current = 0;
                int cycles = 0;

                // Robust extraction using SelectSingleNode if possible, fallback to loop
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

                using (var client = new HttpClient())
                {
                    var json = JsonSerializer.Serialize(payload);
                    var content = new StringContent(json, Encoding.UTF8, "application/json");
                    await client.PostAsync("https://battery-health-checker-d9eac7e3.base44.app/api/apps/6978e65367501f48d9eac7e3/entities/BatteryTest", content);
                }
                
                log("הדו\"ח שוגר למסד בהצלחה!");
                try { File.Delete(tempXml); } catch { }
            }
            catch (Exception ex)
            {
                log("שגיאה במנגנון API הסוללה: " + ex.Message);
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
                    // Only looking for the first physical disk for simplicity
                    string health = disk["HealthStatus"]?.ToString() == "0" ? "Healthy (תקין)" : "Warning (אזהרה)";
                    
                    // Attempt to get wear if supported (mostly NVMe)
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
            catch { }
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
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("defrag", "C: /O /U")
                {
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };
                using (Process p = Process.Start(psi))
                {
                    string output = p.StandardOutput.ReadToEnd();
                    p.WaitForExit();

                    if (p.ExitCode == 0)
                    {
                        log("אופטימיזציית כונן הושלמה בהצלחה ✓");
                    }
                    else
                    {
                        log("אופטימיזציית כונן הסתיימה עם אזהרה (Exit: " + p.ExitCode + ")");
                    }
                    return output;
                }
            }
            catch (Exception ex)
            {
                log("שגיאה באופטימיזציית כונן: " + ex.Message);
                return "Error: " + ex.Message;
            }
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
                        // Fade envelope: smooth attack (10ms) and release (20ms)
                        double envelope = 1.0;
                        int attackSamples = sampleRate / 100; // 10ms
                        int releaseSamples = sampleRate / 50; // 20ms
                        if (i < attackSamples) envelope = (double)i / attackSamples;
                        if (i > samples - releaseSamples) envelope = (double)(samples - i) / releaseSamples;

                        double sample = Math.Sin(2 * Math.PI * freq * i / sampleRate) * 12000 * envelope;
                        short s = (short)sample;
                        BitConverter.GetBytes(s).CopyTo(buffer, offset);       // Left
                        BitConverter.GetBytes(s).CopyTo(buffer, offset + 2);   // Right
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

        public static void PlayStereoTest(bool leftSide)
        {
            // Simple approach: create a temporary wav and play it. 
            // Better: generate byte array and use MemoryStream.
            byte[] wav = CreateBeepWav(440, 1.0, leftSide);
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
                short s = (short)(Math.Sin(2 * Math.PI * freq * i / sampleRate) * 16000);
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
            string[] paths = {
                Path.GetTempPath(),
                @"C:\Windows\Temp",
                @"C:\Windows\Prefetch",
                @"C:\Windows\SoftwareDistribution\Download"
            };

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
            mciSendString("play record.wav wait", null, 0, IntPtr.Zero);
        }

        // CAMERA
        public static void OpenCameraFor30Secs()
        {
            Process.Start(new ProcessStartInfo("microsoft.windows.camera:") { UseShellExecute = true });
        }
        
        public static void CloseCamera()
        {
            RunProcess("taskkill", "/IM WindowsCamera.exe /F");
        }

        // ----------------- STATE MACHINE & SYSTEM -----------------

        public static void SetupAutoResumeAndRestart(int targetPhase, string tech, string serial, string cpu, bool isFullAuto, bool skipHostname)
        {
            // Use a persistent folder (AppData) for the state file
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            string winAutoDir = Path.Combine(appData, "WinAutomator");
            if (!Directory.Exists(winAutoDir)) Directory.CreateDirectory(winAutoDir);
            
            string statePath = Path.Combine(winAutoDir, "WinAutomatorState.txt");
            File.WriteAllText(statePath, $"{tech}|{serial}|{isFullAuto}|{targetPhase}|{cpu}|{skipHostname}");

            // Best way to find the current EXE path in .NET 8
            string exePath = Environment.ProcessPath ?? Process.GetCurrentProcess().MainModule?.FileName ?? "";
            if (string.IsNullOrEmpty(exePath)) return;

            // LAYER 1: Scheduled Task (Highest Privilege, Zero Delay)
            try
            {
                string taskName = "WinAutomator_Resume";
                string trArg = $"\"{exePath}\" --resume";
                
                RunProcess("schtasks", $"/delete /tn \"{taskName}\" /f");
                // Delay 0000:00 means immediate launch
                string schArgs = $"/create /tn \"{taskName}\" /tr \"{trArg}\" /sc onlogon /rl highest /f /delay 0000:00";
                RunProcess("schtasks", schArgs);
            }
            catch { }

            // LAYER 2: Registry RunOnce (Fastest standard Windows trigger)
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", true))
                {
                    if (key != null)
                        key.SetValue("WinAutomator_Resume", $"\"{exePath}\" --resume");
                }
            }
            catch { }

            // LAYER 3: Startup Folder (Legacy Fallback)
            try
            {
                string startupFolder = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
                string batPath = Path.Combine(startupFolder, "WinAutomator_Resume.bat");
                // No timeout, just start
                string batContent = $"@echo off\r\nstart \"\" \"{exePath}\" --resume\r\ndel \"%~f0\"";
                File.WriteAllText(batPath, batContent);
            }
            catch { }

            // Final check: if everything fails, we still have the state file for manual launch
            RunProcess("shutdown", "/r /t 0 /f");
        }

        public static void CleanupAutoResume()
        {
            try
            {
                // Delete task
                RunProcess("schtasks", "/delete /tn \"WinAutomator_Resume\" /f");

                // Cleanup Startup
                string startupFolder = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
                string batPath = Path.Combine(startupFolder, "WinAutomator_Resume.bat");
                if (File.Exists(batPath)) File.Delete(batPath);

                // RunOnce is automatically deleted by Windows, but we ensure it's gone
                try {
                    using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce", true))
                    {
                        if (key != null && key.GetValue("WinAutomator_Resume") != null)
                            key.DeleteValue("WinAutomator_Resume");
                    }
                } catch {}

                // Delete the state file
                string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                string statePath = Path.Combine(appData, "WinAutomator", "WinAutomatorState.txt");
                if (File.Exists(statePath)) File.Delete(statePath);
            }
            catch { }
        }

        private static void RunProcess(string fileName, string arguments)
        {
            ProcessStartInfo psi = new ProcessStartInfo(fileName, arguments) { CreateNoWindow = true, UseShellExecute = false };
            using (Process p = Process.Start(psi)) p.WaitForExit();
        }
    }
}
