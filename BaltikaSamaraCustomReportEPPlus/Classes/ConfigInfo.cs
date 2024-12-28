using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaltikaSamaraCustomReportEPPlus.Classes
{
    public class ConfigInfo
    {
        public DateTime StartDT { get; set; }

        public DateTime EndDT { get; set; }

        public DateTime ReportGeneration { get; set; }

        public string[]? PointsArray { get; set; }

        private string[]? configInfo;
        public ConfigInfo(string patnInlet = Constants.Path.PATH_CONFIGINFO) 
        {
            configInfo = File.ReadAllLines(patnInlet);

            StartDT = Convert.ToDateTime(configInfo[0]);
            EndDT = Convert.ToDateTime(configInfo[1]);
            ReportGeneration = Convert.ToDateTime(configInfo[3]);

            PointsArray = configInfo[2].Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        }
    }
}
