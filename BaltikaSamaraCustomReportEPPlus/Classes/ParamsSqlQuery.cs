using BaltikaSamaraCustomReportEPPlus.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaltikaSamaraCustomReportEPPlus.Classes
{
    public class ParamsSqlQuery
    {
        public string TableName { get; set; } = string.Empty;

        public string MinParam { get; set; } = string.Empty;

        public string MaxParam { get; set; } = string.Empty;

        public string AvgParam { get; set; } = string.Empty;

        public ParamsSqlQuery(Table table) 
        {
            switch (table)
            {
                case 0:
                    TableName = "MinuteData";
                    MinParam = "Min";
                    MaxParam = "Max";
                    AvgParam = "Avg";
                    break;
                default:
                    TableName = "HourData";
                    MinParam = "Min";
                    MaxParam = "Max";
                    AvgParam = "Avg";
                    break;
            }
        }
    }
}
