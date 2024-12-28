using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaltikaSamaraCustomReportEPPlus.Classes
{
    public class Document
    {
        const int START_REPORT_ROW = 6;

        static public string[] cells = {
            "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA"
        };

        static public int currentLimitsRow = START_REPORT_ROW;
        static public int currentValuesRow = START_REPORT_ROW;
    }
}
