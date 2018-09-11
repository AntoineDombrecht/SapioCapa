using System;
using System.IO;

namespace MailConnector
{
    public class Output
    {
        private static string path = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName +
            @System.Configuration.ConfigurationManager.AppSettings["Local_DebugPath"];
        private static string dateTimeFormat = System.Configuration.ConfigurationManager.AppSettings["Program_DateTimeFormat"];
        private static string repTemp = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName +
            @System.Configuration.ConfigurationManager.AppSettings["Local_TempPath"];
        private string log;
        private string state;
        private string date;
        private static DateTime begDate;
        private static int tempSize = Convert.ToInt32(System.Configuration.ConfigurationManager.AppSettings["Program_TempSize"]);
        public Output(string _log = null, string _state = null, string _date = null)
        {
            // Initialize parameters (optional)
            log = _log;
            state = _state;
            date = _date;
        }
        // Write the logs into the output folder
        public static void WriteInFile(string state, string log)
        {
            string text = DateTime.Now.ToString(dateTimeFormat) + " " + state + " : " + log;
            text += Environment.NewLine;
            File.AppendAllText(path, text);
        }
        // Initialize the log file
        public static void Initialize()
        {
            // Clear the previous logs
            File.WriteAllText(path, String.Empty);
            // Title
            string text = " ==================== " + Environment.NewLine;
                   text += " ===  Output Logs === " + Environment.NewLine;
                   text += " ==================== " + Environment.NewLine;
                   text += Environment.NewLine;
            File.AppendAllText(path, text);
            Output.WriteInFile("[OK]", "Lancement du programme");
            begDate = DateTime.Now;
        }
        // clean the temp folder
        public static void cleanTempDir()
        {
            System.IO.DirectoryInfo di = new DirectoryInfo(repTemp);
            foreach (FileInfo file in di.GetFiles())
            {
                if(file.Length >= tempSize || tempSize == 0)
                    file.Delete();
            }
            foreach (DirectoryInfo dir in di.GetDirectories())
            {
                dir.Delete(true);
            }
        }
        // save all the temp folder into the FTP repository
        public static void saveInFTP()
        {
            // Save in the corresponding FTP server file
            FTPClient ftpClient = new FTPClient();

            System.IO.DirectoryInfo di = new DirectoryInfo(repTemp);
            foreach (DirectoryInfo dir in di.GetDirectories())
            {
                ftpClient.saveInFile(dir.Name);
            }
        }
        public static void End()
        {
            DateTime endDate = DateTime.Now;
            System.TimeSpan diff = endDate.Subtract(begDate);
            Output.WriteInFile("[INFO]", "Fermeture du programme. Temps d'execution : "+diff.ToString());
        }
    }
}