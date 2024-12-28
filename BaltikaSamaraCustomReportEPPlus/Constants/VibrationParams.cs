using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaltikaSamaraCustomReportEPPlus.Constants
{
    public static class VibrationParams
    {
        //формат массивов (номера каналов в БД из Rapid SCADA):

        public static int[] millstar_2 = { 110, 113, 107, 210, 213, 207, 310, 313, 307, 510, 513, 507, 610, 613, 607, 710, 713, 707, 810, 813, 807, 410, 413, 407 };

        public static int[] millstar_3 = { 910, 913, 907, 1010, 1013, 1007, 1110, 1113, 1107, 1310, 1313, 1307, 1410, 1413, 1407, 1510, 1513, 1507, 1610, 1613, 1607, 1210, 1213, 1207 };

        public static List<int[]> list = new List<int[]>() { millstar_2, millstar_3 };
    }
}
