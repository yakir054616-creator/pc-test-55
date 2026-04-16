using System;
using System.Windows.Forms;
using System.Security.Principal;
using System.IO;

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
    }
}
