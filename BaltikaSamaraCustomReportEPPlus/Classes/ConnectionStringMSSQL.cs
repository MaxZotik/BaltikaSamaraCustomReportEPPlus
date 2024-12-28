using BaltikaSamaraCustomReportEPPlus.Interfaces;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BaltikaSamaraCustomReportEPPlus.Classes
{
    public class ConnectionStringMSSQL : IConnectDb
    {
        private string path = "";
        public string ConnectionString { get; set; } = "";

        public ConnectionStringMSSQL(string pathInlet = Constants.Path.PATH_DB_INFO)
        {
            path = pathInlet;
            InitConnectionString();
        }

        private void InitConnectionString()
        {
            string[] read_result_array = File.ReadAllLines(path, Encoding.Default);

            ConnectionString = $@"Data Source={read_result_array[0]};Initial Catalog={read_result_array[1]};User ID={read_result_array[2]};Password={read_result_array[3]}";
        }
    }
}
