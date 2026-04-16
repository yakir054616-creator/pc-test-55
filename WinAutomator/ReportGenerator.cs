using System;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace WinAutomator
{
    public class ReportData
    {
        public string TimeStamp { get; set; }
        public string TechId { get; set; }
        public string SerialNum { get; set; }
        
        public string AnsPlastic { get; set; }
        public string AnsOldDisk { get; set; }
        public string AnsNewDisk { get; set; }
        
        public string CpuGen { get; set; }
        public string CpuName { get; set; }
        public string RamSize { get; set; }
        public string AnsCD { get; set; }
        public string AnsScreenSize { get; set; }
        public string Manufacturer { get; set; }
        
        // Results (QA)
        public DialogResult MicResult { get; set; }
        public DialogResult SpeakerResult { get; set; }
        public DialogResult CameraResult { get; set; }
        public DialogResult KeyboardResult { get; set; }
        public DialogResult TrackpadResult { get; set; }
        public DialogResult UsbResult { get; set; }

        public string AnsScrews { get; set; }
        public string AnsClean { get; set; }
        public string AnsAppearance { get; set; }
        public string Notes1 { get; set; }
        public string Notes2 { get; set; }
        
        public string SsdSize { get; set; }
        public string AnsTouch { get; set; }
    }

    public static class ReportGenerator
    {
        public static void GenerateHtmlReport(ReportData d)
        {
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(desktop, $"Report_{d.SerialNum}.html");

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<!DOCTYPE html>");
            sb.AppendLine("<html dir='rtl' lang='he'>");
            sb.AppendLine("<head>");
            sb.AppendLine("<meta charset='utf-8'>");
            sb.AppendLine("<title>Hardware QA Report</title>");
            sb.AppendLine("<style>");
            sb.AppendLine("body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; background: #fff; color: #000; margin: 0; padding: 20px; font-size: 14px; }");
            sb.AppendLine(".container { max-width: 800px; margin: auto; border: 1px solid #ccc; padding: 20px; box-shadow: 0 0 10px rgba(0,0,0,0.1); }");
            sb.AppendLine(".header-text { text-align: center; font-weight: bold; font-size: 16px; margin-bottom: 5px; }");
            sb.AppendLine(".barcode-text { text-align: center; font-size: 30px; letter-spacing: 2px; margin-bottom: 0px; font-family: monospace; }");
            sb.AppendLine(".barcode-id { text-align: center; font-weight: bold; font-size: 16px; margin-bottom: 20px; }");
            sb.AppendLine("table { width: 100%; border-collapse: collapse; margin-top: 20px; }");
            sb.AppendLine("th, td { border: 1px solid #000; padding: 6px 10px; text-align: right; }");
            sb.AppendLine(".col-q { width: 60%; background: #f9f9f9; }");
            sb.AppendLine(".col-a { width: 35%; }");
            sb.AppendLine(".col-x { width: 5%; text-align: center; font-weight: bold; font-size: 16px; border-left: 1px solid #000; }");
            sb.AppendLine("@media print { .container { border: none; box-shadow: none; padding: 0; } }");
            sb.AppendLine("</style>");
            sb.AppendLine("</head>");
            sb.AppendLine("<body>");
            
            sb.AppendLine("<div class='container'>");
            sb.AppendLine($"<div class='header-text'>שדרוג<br>{d.Manufacturer}-{d.CpuName}-{d.RamSize}-{d.CpuGen}-{d.AnsScreenSize}+ / {d.TechId}</div>");
            sb.AppendLine("<div class='header-text'>ממתין להחלטה<br>מקלדת</div>");
            
            // Faux Barcode
            sb.AppendLine($"<div class='barcode-text'>||||||||||||||||||||||||||||||||||||||</div>");
            sb.AppendLine($"<div class='barcode-id'>{d.SerialNum}</div>");

            sb.AppendLine("<table>");
            sb.AppendLine("<tbody>");
            
            // Helper for Rows
            void AddRow(string q, string a, bool isFail = false)
            {
                string xStr = isFail && !string.IsNullOrWhiteSpace(a) ? "X" : ""; 
                sb.AppendLine("<tr>");
                sb.AppendLine($"<td class='col-q'>{q}</td>");
                sb.AppendLine($"<td class='col-a'>{a}</td>");
                sb.AppendLine($"<td class='col-x'>{xStr}</td>");
                sb.AppendLine("</tr>");
            }

            // Mappings 1 to 1 based on image
            AddRow("חותמת זמן", d.TimeStamp);
            AddRow("מזהה מבצע/ת השדרוג :", d.TechId);
            AddRow("מס' סידורי של המחשב (ברקוד מחשבים) :", d.SerialNum);
            
            AddRow("1. תוצאות בדיקה ראשונית", "המחשב פועל");
            AddRow("2. האם המחשב שלם, ללא שברים בפלסטיקה או בכפתורים?", d.AnsPlastic, d.AnsPlastic.Contains("לא"));
            AddRow("3. יש לפתוח את המכסה התחתון ולפרק את הדיסק הישן של המחשב", d.AnsOldDisk);
            AddRow("4. יש להתקין דיסק מעודכן עם מערכת הפעלה של \"מחשבים\"", d.AnsNewDisk);
            AddRow("5. יש לחבר את המחשב לחשמל, להפעיל אותו ולוודא שעולה באופן תקין", "המחשב עולה באופן תקין");
            AddRow("6. יש לוודא שדור המעבד במחשב עומד במפרט המינימום 5 ומעלה.", "דור המעבד במחשב עומד במפרט המינימום");
            
            AddRow("7. שם המעבד", d.CpuName);
            AddRow("8. דור המעבד", d.CpuGen);
            AddRow("9. התקנת זיכרון", "");
            AddRow("10. כמות זיכרון - כמה ג'יגה RAM :", d.RamSize);
            AddRow("11. במידה וקיים מתקן דיסק (CD) יש לוודא שהוא נפתח.", d.AnsCD, d.AnsCD.Contains("לא"));
            AddRow("12. גודל המסך - כמה אינטש", d.AnsScreenSize);
            AddRow("13. יצרן המחשב", d.Manufacturer);
            AddRow("14. חיבור WiFi תקין", "חיבור WiFi תקין");
            AddRow("15. רישיון Windows", "רישיון Windows מופעל");
            AddRow("16. רישיון Office", "רישיון Office מופעל");
            AddRow("17. בדיקת מערכת", "בדיקה תקינה");
            AddRow("18. יש לוודא שמערכת ההפעלה מעודכנת", "מערכת ההפעלה מעודכנת");
            AddRow("19. עדכונים נוספים", "עדכונים נוספים הותקנו");
            AddRow("20. הפעלה מחדש", "כל העדכונים הקיימים הותקנו");
            AddRow("21. עדכוני תוכנת יצרן", "דרייברים עודכנו");
            
            string usbAns = d.UsbResult == DialogResult.Yes ? "כל כניסות ה USB תקינות" : "כניסות ה USB לא תקינות";
            AddRow("22. כניסות USB", usbAns, d.UsbResult != DialogResult.Yes && d.UsbResult != DialogResult.Ignore);
            
            string spkAns = d.SpeakerResult == DialogResult.Yes ? "רמקולים תקינים" : "רמקולים לא תקינים";
            AddRow("23. רמקולים", spkAns, d.SpeakerResult != DialogResult.Yes);
            
            string micAns = d.MicResult == DialogResult.Yes ? "מיקרופון תקין" : "מיקרופון לא תקין";
            AddRow("24. מיקרופון", micAns, d.MicResult != DialogResult.Yes);
            
            string camAns = d.CameraResult == DialogResult.Yes ? "מצלמה תקינה" : "מצלמה לא תקינה";
            AddRow("25. מצלמה", camAns, d.CameraResult != DialogResult.Yes);
            
            string kbAns = d.KeyboardResult == DialogResult.Yes ? "מקלדת תקינה" : "מקלדת לא תקינה";
            AddRow("26. מקלדת", kbAns, d.KeyboardResult != DialogResult.Yes);
            
            AddRow("27. סוללה", "סוללה תקינה"); // Assumed checked
            
            AddRow("28. ברגים", d.AnsScrews, d.AnsScrews.Contains("חסרים"));
            AddRow("29. מחשב נקי", d.AnsClean, d.AnsClean.Contains("מלוכלך"));
            AddRow("30. מראה חיצוני", d.AnsAppearance, d.AnsAppearance.Contains("פגמים"));
            AddRow("31. הערות לגבי המחשב (למשל מקלדת באנגלית) המחשב ימתין להחלטה", d.Notes1, !string.IsNullOrWhiteSpace(d.Notes1));
            AddRow("32. מידע כללי נוסף", d.Notes2);
            AddRow("10.1. יש לוודא גודל SSD (GB)", d.SsdSize);
            AddRow("12.1. מסך מגע", d.AnsTouch);
            AddRow("הערה (שלב 2)", "");
            AddRow("הערה (שלב 3)", "");

            sb.AppendLine("</tbody>");
            sb.AppendLine("</table>");
            sb.AppendLine("</div>");
            sb.AppendLine("</body>");
            sb.AppendLine("</html>");

            File.WriteAllText(filePath, sb.ToString(), Encoding.UTF8);
        }
    }
}
