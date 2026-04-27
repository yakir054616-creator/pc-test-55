using System;
using System.Windows.Forms;
using System.Security.Principal;
using System.IO;
using System.Diagnostics;

namespace WinAutomator
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            // Use a mutex to ensure only one instance runs, even if multiple startup triggers (Task, Registry, Startup) fire at once.
            using (var mutex = new System.Threading.Mutex(true, "Global\\WinAutomator_Singleton_7733", out bool createdNew))
            {
                if (!createdNew) return; 

                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                
                if (!IsAdministrator())
                {
                    MessageBox.Show("שגיאה! עליך להריץ כלי זה כמנהל (Run as administrator) עבור החלת פקודות למשאבי מערכת.", 
                        "הרשאות חסרות", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // === Self-copy to Desktop ===
                // If running from a shared/network folder, copy the EXE to the Desktop and relaunch from there.
                // Skip this logic on --resume (post-reboot) since we're already on the desktop.
                bool isResuming = args.Length > 0 && args[0].ToLower() == "--resume";
                if (!isResuming && ShouldCopyToDesktop(out string desktopExePath))
                {
                    try
                    {
                        string currentExe = Environment.ProcessPath ?? System.Reflection.Assembly.GetExecutingAssembly().Location;
                        
                        // Copy the single EXE to the Desktop
                        File.Copy(currentExe, desktopExePath, overwrite: true);

                        // Relaunch from the desktop copy with admin privileges
                        var psi = new ProcessStartInfo
                        {
                            FileName = desktopExePath,
                            UseShellExecute = true,
                            Verb = "runas",
                            Arguments = string.Join(" ", args)
                        };
                        Process.Start(psi);
                        return; // Exit this instance
                    }
                    catch (Exception ex)
                    {
                        // If copy fails, just continue running from current location
                        MessageBox.Show(
                            $"לא הצלחתי להעתיק לשולחן העבודה, ממשיך מהמיקום הנוכחי.\n{ex.Message}",
                            "אזהרה", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

                int currentPhase = 1;
                string tech = "";
                string serial = "";
                string cpu = "";
                bool isFullAuto = true;
                bool skipHostname = false;
                string manufacturer = "";
                
                if (args.Length > 0 && args[0].ToLower() == "--resume")
                {
                    string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                    string statePath = Path.Combine(appData, "WinAutomator", "WinAutomatorState.txt");
                    
                    if (File.Exists(statePath))
                    {
                        try {
                            string[] parts = File.ReadAllText(statePath).Split('|');
                            if (parts.Length >= 5)
                            {
                                tech = parts[0];
                                serial = parts[1];
                                bool.TryParse(parts[2], out isFullAuto);
                                int.TryParse(parts[3], out currentPhase);
                                cpu = parts[4];
                                if (parts.Length >= 6) bool.TryParse(parts[5], out skipHostname);
                                if (parts.Length >= 7) manufacturer = parts[6];
                            }
                        } catch { currentPhase = 1; }
                    }
                }
                else 
                {
                    // Startup warning - only show on fresh launch (not after a reboot/resume)
                    MessageBox.Show(
                        "שים לב!\nהתוכנה לא מהווה תחליף לעדכוני וינדוס או יצרן, חובה לעשות אותם בנוסף!!",
                        "הוראות שימוש", 
                        MessageBoxButtons.OK, 
                        MessageBoxIcon.Warning, 
                        MessageBoxDefaultButton.Button1, 
                        MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                }

                Application.Run(new MainForm(currentPhase, tech, serial, cpu, isFullAuto, skipHostname, manufacturer));
            }
        }

        public static bool IsAdministrator()
        {
            try {
                WindowsIdentity identity = WindowsIdentity.GetCurrent();
                WindowsPrincipal principal = new WindowsPrincipal(identity);
                return principal.IsInRole(WindowsBuiltInRole.Administrator);
            } catch { return false; }
        }

        /// <summary>
        /// Checks if the current exe is running from outside the Desktop.
        /// If so, returns true and outputs the target desktop path (Desktop\PROJECT.exe).
        /// </summary>
        private static bool ShouldCopyToDesktop(out string desktopExePath)
        {
            desktopExePath = "";
            try
            {
                string currentExe = Environment.ProcessPath ?? System.Reflection.Assembly.GetExecutingAssembly().Location;
                string currentDir = Path.GetDirectoryName(currentExe) ?? "";
                string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                // Target: Desktop\PROJECT.exe
                desktopExePath = Path.Combine(desktop, "PROJECT.exe");

                // Already running from the desktop? No need to copy.
                if (currentDir.StartsWith(desktop, StringComparison.OrdinalIgnoreCase))
                    return false;

                // Already running from LocalAppData? (post-resume scenario)
                string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
                if (currentDir.StartsWith(localAppData, StringComparison.OrdinalIgnoreCase))
                    return false;

                return true;
            }
            catch { return false; }
        }
    }
}
