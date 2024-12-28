using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaltikaSamaraCustomReportEPPlus.Interfaces
{
    public interface IConnectDb
    {
        string ConnectionString { get; set; }
    }
}
