using BaltikaSamaraCustomReportEPPlus.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaltikaSamaraCustomReportEPPlus.Classes
{
    public class Logger
    {
        public static async Task LoggingAsync(string path, string message)
        {
            CheckFile(path);

            await File.AppendAllTextAsync(path, DateTime.Now.ToString() + " - " + message + Environment.NewLine, Encoding.Default);
        }

        public static async Task WriteStatusAsync(string path, Status status)
        {
            using (StreamWriter sw = new StreamWriter(path, false, Encoding.Default))
            {
                await sw.WriteLineAsync(((int)status).ToString());
            }
        }

        private static void CheckFile(string path)
        {
            if (File.Exists(path) == false)
            {
                File.Create(path).Close();
            }
        }
    }
}
