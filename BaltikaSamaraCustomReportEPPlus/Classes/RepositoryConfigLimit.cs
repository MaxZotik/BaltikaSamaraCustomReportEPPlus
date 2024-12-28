using BaltikaSamaraCustomReportEPPlus.Interfaces;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BaltikaSamaraCustomReportEPPlus.Constants;

namespace BaltikaSamaraCustomReportEPPlus.Classes
{
    public class RepositoryConfigLimit : IRepository
    {
        private readonly string _connection;
        private DataTable? TableDatas;
        public List<double[,]> LimitsValues { get; set; }

        public RepositoryConfigLimit(IConnectDb connectionString) 
        {
            _connection = connectionString.ConnectionString;
            LimitsValues = new List<double[,]>();
            GetDataDbAsync();
            GetLimitsValues();
        }

        private async void GetDataDbAsync()
        {
            string selectQuery = $"SELECT [CnlNum], [Limit] FROM ConfigLimits";

            try
            {
                using (SqlConnection connection = new SqlConnection(_connection))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter(selectQuery, connection);
                    DataSet ds = new DataSet();
                    adapter.Fill(ds);

                    TableDatas = ds.Tables[0];
                }
            }
            catch (Exception ex)
            {
                await Logger.LoggingAsync(Constants.Path.PATH_LOG, ex.ToString());
            }
        }

        private void GetLimitsValues()
        {
            for (int i = 0; i < Limits.listLimits.Count; i++)
            {
                double[,] temp = new double[Limits.listLimits[i].GetLength(0), Limits.listLimits[i].GetLength(1)];

                LimitsValues.Add(temp);

                for (int j = 0; j < Limits.listLimits[i].GetLength(0); j++)
                {
                    for (int k = 0; k < Limits.listLimits[i].GetLength(1); k++)
                    {
                        LimitsValues[i][j, k] = double.Parse(TableDatas.Select($"CnlNum = {Limits.listLimits[i][j, k]}").First().ItemArray[1].ToString());
                    }
                }
            }
        }
    }
}
