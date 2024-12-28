using BaltikaSamaraCustomReportEPPlus.Classes;
using BaltikaSamaraCustomReportEPPlus.Constants;
using BaltikaSamaraCustomReportEPPlus.Enums;
using BaltikaSamaraCustomReportEPPlus.Interfaces;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;

namespace BaltikaSamaraCustomReportEPPlus
{
    internal class StartApp
    {
        const int MIN_POINT_COUNT = 5;
        const int STAT_TYPE_COUNT = 3;
        const int SECTION_INTERVAL_ROW_COUNT = 5;

        public static async void StartAsync()
        {
            try
            {
                IConnectDb connectionString = new ConnectionStringMSSQL();

                await Logger.WriteStatusAsync(Constants.Path.PATH_CHECK, Status.Process);

                ConfigInfo configInfo = new ConfigInfo();
                RepositoryConfigLimit repositoryConfigLimit = new RepositoryConfigLimit(connectionString);

                string sqlStartDT = ConvertDataSql(configInfo.StartDT);
                string sqlEndDT = ConvertDataSql(configInfo.EndDT);

                SqlConnection conn = new SqlConnection(connectionString.ConnectionString);
                SqlCommand sqlCommand = new SqlCommand();

                //Теперь необходимо выбрать таблицу и забрать данные, после чего можно размещать данные в Excel + постройка графиков

                //Если в указанный интервал времени есть 5 или более точек в часовой таблице, то делаем по ней, иначе проверка на наличие 5 или более точек в таблице с темпом измерений
                conn.Open();

                //string tableName = string.Empty, maxParam = string.Empty, minParam = string.Empty, avgParam = string.Empty;
                int hourCount = 0;
                int cnlCount = 0;

                for (int i = 0; i < VibrationParams.list.Count; i++)
                {
                    for (int j = 0; j < VibrationParams.list[i].Length; j++)
                    {
                        string hourRowCountCheck = "SELECT COUNT(*) FROM HourData WHERE CnlNum = " + VibrationParams.list[i][j] + " AND DateTime BETWEEN CONVERT(datetime, '" + sqlStartDT + "', 20) AND CONVERT(datetime, '" + sqlEndDT + "', 20)";
                        sqlCommand = new SqlCommand(hourRowCountCheck, conn);
                        int hourTemp = Convert.ToInt32(sqlCommand.ExecuteScalar());
                        if (hourTemp > hourCount)
                        {
                            hourCount = hourTemp;
                        }

                        string cnlRowCountCheck = "SELECT COUNT(*) FROM MinuteData WHERE CnlNum = " + VibrationParams.list[i][j] + " AND DateTime BETWEEN CONVERT(datetime, '" + sqlStartDT + "', 20) AND CONVERT(datetime, '" + sqlEndDT + "', 20)";
                        sqlCommand = new SqlCommand(cnlRowCountCheck, conn);
                        int cnlTemp = Convert.ToInt32(sqlCommand.ExecuteScalar());
                        if (cnlTemp > cnlCount)
                        {
                            cnlCount = cnlTemp;
                        }
                    }
                }
                conn.Close();

                ParamsSqlQuery paramsSqlQuery;

                if (hourCount >= MIN_POINT_COUNT)
                {
                    paramsSqlQuery = new ParamsSqlQuery(Table.HourData);
                }
                else
                {
                    if (cnlCount >= MIN_POINT_COUNT)
                    {
                        paramsSqlQuery = new ParamsSqlQuery(Table.MinuteData);
                    }
                    else
                    {
                        throw new Exception("Отсутствуют данные за указанный период (" + configInfo.StartDT + " - " + configInfo.EndDT + ")");
                    }
                }

                List<double[,]> paramsValueList = new List<double[,]>();

                GetVibrationValues(conn, STAT_TYPE_COUNT, paramsValueList, sqlCommand, paramsSqlQuery, sqlStartDT, sqlEndDT);

                //Заполнение документа Excel
                FileInfo fileInfo = new FileInfo(Constants.Path.PATH_REPORT_SCHEME);

                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                ExcelPackage excel = new ExcelPackage(fileInfo);
                
                var worksheetsOne = excel.Workbook.Worksheets[0];

                worksheetsOne.Cells["C2"].Value = configInfo.ReportGeneration.ToString();
                worksheetsOne.Cells["E4"].Value = configInfo.StartDT.ToString();
                worksheetsOne.Cells["H4"].Value = configInfo.EndDT.ToString();

                #region

                var worksheetsTwo = excel.Workbook.Worksheets[1];

                for (int i = 0; i < repositoryConfigLimit.LimitsValues.Count; i++)
                {
                    for (int j = 0; j < repositoryConfigLimit.LimitsValues[i].GetLength(0); j++)
                    {
                        for (int k = 0; k < repositoryConfigLimit.LimitsValues[i].GetLength(1); k++)
                        {
                            string index = Document.cells[k] + Document.currentLimitsRow;
                            worksheetsTwo.Cells[index].Value = Math.Round(repositoryConfigLimit.LimitsValues[i][j, k], 2);
                        }
                        Document.currentLimitsRow += 1;
                    }
                    Document.currentLimitsRow += SECTION_INTERVAL_ROW_COUNT;
                }


                var worksheetsThree = excel.Workbook.Worksheets[2];

                int tempIndex = 0;
                for (int i = 0; i < paramsValueList.Count; i++)
                {
                    for (int j = 0; j < paramsValueList[i].GetLength(1); j++)
                    {
                        for (int k = 0; k < paramsValueList[i].GetLength(0); k++)
                        {
                            bool check = false;
                            int length = configInfo.PointsArray.Length;
                            string index = Document.cells[k] + Document.currentValuesRow;

                            for (int n = 0; n < length; n++)
                            {
                                if (Convert.ToInt32(configInfo.PointsArray[n]) == tempIndex + k / 3)
                                {
                                    check = true;
                                }
                            }

                            if (check == true)
                            {
                                worksheetsThree.Cells[index].Value = Math.Round(paramsValueList[i][k, j], 2);
                            }
                            else
                            {
                                worksheetsThree.Cells[index].Value = "---";
                            }

                        }

                        if (j == paramsValueList[i].GetLength(1) - 1)
                        {
                            tempIndex += paramsValueList[i].GetLength(0) / STAT_TYPE_COUNT;
                        }

                        Document.currentValuesRow += 1;
                    }

                    Document.currentValuesRow += SECTION_INTERVAL_ROW_COUNT;

                }

                string savePath =
                    Constants.Path.PATH_SAVE +
                    (configInfo.ReportGeneration.Day >= 10 ? configInfo.ReportGeneration.Day.ToString() : "0" + configInfo.ReportGeneration.Day.ToString()) + "." +
                    (configInfo.ReportGeneration.Month >= 10 ? configInfo.ReportGeneration.Month.ToString() : "0" + configInfo.ReportGeneration.Month.ToString()) + "." +
                    configInfo.ReportGeneration.Year + "_" +
                    (configInfo.ReportGeneration.Hour >= 10 ? configInfo.ReportGeneration.Hour.ToString() : "0" + configInfo.ReportGeneration.Hour.ToString()) + "." +
                    (configInfo.ReportGeneration.Minute >= 10 ? configInfo.ReportGeneration.Minute.ToString() : "0" + configInfo.ReportGeneration.Minute.ToString()) + "." +
                    (configInfo.ReportGeneration.Second >= 10 ? configInfo.ReportGeneration.Second.ToString() : "0" + configInfo.ReportGeneration.Second.ToString()) +
                    ".xlsx";

                excel.SaveAs(savePath);
                //excelWorkbook.Close();

                #endregion

                await Logger.WriteStatusAsync(Constants.Path.PATH_CHECK, Status.Ready);
                await Logger.LoggingAsync(Constants.Path.PATH_LOG, "Формирование отчёта успешно завершено");
                //KillExcelProcess();

            }
            catch (Exception ex)
            {
                await Logger.WriteStatusAsync(Constants.Path.PATH_CHECK, Status.Error);
                await Logger.LoggingAsync(Constants.Path.PATH_LOG, ex.ToString());
                //KillExcelProcess();
            }
        }

        public static void GetVibrationValues(SqlConnection conn, int statTypeCount, List<double[,]> paramsValueList, SqlCommand sqlCommand, ParamsSqlQuery paramsSqlQuery, string sqlStartDT, string sqlEndDT)
        {
            conn.Open();
            for (int i = 0; i < VibrationParams.list.Count; i++)
            {
                double[,] temp = new double[VibrationParams.list[i].Length, statTypeCount];
                paramsValueList.Add(temp);
                for (int j = 0; j < VibrationParams.list[i].Length; j++)
                {
                    double min, max, avg;

                    string minQuery = "SELECT ISNULL ((SELECT MIN(" + paramsSqlQuery.MinParam + ") FROM " + paramsSqlQuery.TableName + " WHERE CnlNum = " + VibrationParams.list[i][j] + " AND DateTime BETWEEN CONVERT(datetime, '" + sqlStartDT + "', 20) AND CONVERT(datetime, '" + sqlEndDT + "', 20)), 0)";
                    sqlCommand = new SqlCommand(minQuery, conn);
                    min = Convert.ToDouble(sqlCommand.ExecuteScalar());
                    Math.Round(min, 2);
                    paramsValueList[i][j, 0] = min;

                    string maxQuery = "SELECT ISNULL ((SELECT MAX(" + paramsSqlQuery.MaxParam + ") FROM " + paramsSqlQuery.TableName + " WHERE CnlNum = " + VibrationParams.list[i][j] + " AND DateTime BETWEEN CONVERT(datetime, '" + sqlStartDT + "', 20) AND CONVERT(datetime, '" + sqlEndDT + "', 20)), 0)";
                    sqlCommand = new SqlCommand(maxQuery, conn);
                    max = Convert.ToDouble(sqlCommand.ExecuteScalar());
                    Math.Round(max, 2);
                    paramsValueList[i][j, 1] = max;

                    string avgQuery = "SELECT ISNULL ((SELECT AVG(" + paramsSqlQuery.AvgParam + ") FROM " + paramsSqlQuery.TableName + " WHERE CnlNum = " + VibrationParams.list[i][j] + " AND DateTime BETWEEN CONVERT(datetime, '" + sqlStartDT + "', 20) AND CONVERT(datetime, '" + sqlEndDT + "', 20)), 0)";
                    sqlCommand = new SqlCommand(avgQuery, conn);
                    avg = Convert.ToDouble(sqlCommand.ExecuteScalar());
                    Math.Round(avg, 2);
                    paramsValueList[i][j, 2] = avg;
                }
            }
            conn.Close();
        }

        private static void KillExcelProcess()
        {
            try
            {
                Process[] proc1 = Process.GetProcessesByName("EXCEL");
                proc1[0].Kill();
            }
            catch
            {

            }
        }

        private static string ConvertDataSql(DateTime dateTime)
        {

            string sqlStartDT =
                   dateTime.Year.ToString() + "-" +
                   (dateTime.Month >= 10 ? dateTime.Month.ToString() : "0" + dateTime.Month.ToString()) + "-" +
                   (dateTime.Day >= 10 ? dateTime.Day.ToString() : "0" + dateTime.Day.ToString()) + " " +
                   (dateTime.Hour >= 10 ? dateTime.Hour.ToString() : "0" + dateTime.Hour.ToString()) + ":" +
                   (dateTime.Minute >= 10 ? dateTime.Minute.ToString() : "0" + dateTime.Minute.ToString()) + ":" +
                   (dateTime.Second >= 10 ? dateTime.Second.ToString() : "0" + dateTime.Second.ToString());

            return sqlStartDT;
        }
    }
}
