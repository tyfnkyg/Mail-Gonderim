using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailGonderme
{
    class Log
    {
        public static void LogMessageToFile(string message)
        {
            if ((ConfigurationManager.AppSettings?["Debug"] ?? "") == "1")
            {
                Console.WriteLine(DateTime.Now.ToString() + " - " + message);
            }

            string logpath = Path.Combine(Environment.CurrentDirectory, "logs");

            if (!Directory.Exists(logpath))
            {
                Directory.CreateDirectory(logpath);
            }

            System.IO.StreamWriter sw = System.IO.File.AppendText(logpath + $"\\{DateTime.Now.ToString("ddMMyyyy")}_log.txt");

            try
            {
                string logLine = System.String.Format(@"{0:G}: {1}.", System.DateTime.Now, message);
                sw.WriteLine(logLine);
            }
            finally
            {
                sw.Close();
            }
        }
    }

}
