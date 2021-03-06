﻿namespace ExcelToDB
{
    using Microsoft.Office.Interop.Excel;
    using MySql.Data.MySqlClient;
    using System;
    using System.Data;
    using System.Data.SqlClient;
    using System.Diagnostics;
    using System.IO;
    using System.Reflection;
    using System.Text;
    using DataTable = System.Data.DataTable;
    using Excel = Microsoft.Office.Interop.Excel;

    /// <summary>
    /// Defines the <see cref="ExcelHelper" />.
    /// </summary>
    public class ExcelHelper
    {
        /// <summary>
        /// The XLSSavesaCSV.
        /// </summary>
        /// <param name="FilePath">The FilePath<see cref="string"/>.</param>
        /// <returns>The <see cref="string"/>.</returns>
        public static string XLSSavesaCSV(string FilePath)
        {
            QuertExcel();
            string new_file_path = "";
            try
            {
                Excel.Application excelApplication = new Excel.ApplicationClass();
                Excel.Workbooks excelWorkBooks = excelApplication.Workbooks;
                Excel.Workbook excelWorkBook = ((Excel.Workbook)excelWorkBooks.Open(FilePath, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value));
                Excel.Worksheet excelWorkSheet = (Excel.Worksheet)excelWorkBook.Worksheets[1];
                excelApplication.Visible = false;
                excelApplication.DisplayAlerts = false;
                string extension = Path.GetExtension(FilePath);
                new_file_path = FilePath.Replace(extension, ".csv");

                ////避免重复创建
                //if (File.Exists(new_file_path))
                //{
                //    DeleteFile(new_file_path);
                //}

                excelWorkBook.SaveAs(new_file_path, XlFileFormat.xlCSV, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                QuertExcel();
            }
            catch (Exception exc)
            {
                throw new Exception(exc.Message);
            }
            return new_file_path;
        }

        /// <summary>
        /// The OpenCSV.
        /// </summary>
        /// <param name="file_path">The file_path<see cref="string"/>.</param>
        /// <returns>The <see cref="DataTable"/>.</returns>
        public static DataTable OpenCSV(string file_path)
        {
            DataTable dt = new DataTable();
            try
            {
                Encoding encoding = Encoding.Default;
                using (FileStream fs = new FileStream(file_path, FileMode.Open, FileAccess.Read))
                {
                    using (StreamReader sr = new StreamReader(fs, encoding))
                    {
                        string strLine = "";
                        string[] aryLine = null;
                        string[] tableHead = null;
                        int columnCount = 0;
                        bool IsFirst = true;
                        //逐行读取CSV中的数据
                        while ((strLine = sr.ReadLine()) != null)
                        {
                            if (IsFirst == true)
                            {
                                tableHead = strLine.Split(',');
                                IsFirst = false;
                                columnCount = tableHead.Length;
                                for (int i = 0; i < columnCount; i++)
                                {

                                    DataColumn dc = new DataColumn(tableHead[i]);
                                    dt.Columns.Add(dc);
                                }
                            }
                            else
                            {
                                if (!String.IsNullOrEmpty(strLine))
                                {
                                    aryLine = strLine.Split(',');
                                    DataRow dr = dt.NewRow();
                                    for (int j = 0; j < columnCount; j++)
                                    {
                                        dr[j] = aryLine[j];
                                    }
                                    dt.Rows.Add(dr);
                                }
                            }
                        }
                        if (aryLine != null && aryLine.Length > 0)
                        {
                            dt.DefaultView.Sort = tableHead[0] + " " + "asc";
                        }
                    }
                }
                return dt;
            }
            catch (Exception ex)
            {

            }

            return null;
        }

        /// <summary>
        /// 转换成和数据库一致的datatable
        /// </summary>
        /// <param name="tb"></param>
        /// <returns></returns>
        private DataTable ConvertDatatTableForDB(DataTable tb)
        {
            DataTable dt = new DataTable();
            try
            {
                dt = tb.Clone();
                foreach (DataColumn col in dt.Columns)
                {
                    switch (col.ColumnName)
                    {
                        case "data_value":
                            col.DataType = typeof(string);
                            break;
                    }
                }

                foreach (DataRow row in tb.Rows)
                {
                    DataRow new_row = dt.NewRow();
                    new_row["data_value"] = row["data_value"];

                    dt.Rows.Add(new_row);
                }

                return dt;
            }
            catch (Exception ex)
            {

            }

            return null;
        }


        public static int SqlBulkCopyInsert(string conStr, string csv_path)
        {
            int result = 0;
            try
            {
                using (MySqlConnection con = new MySqlConnection(conStr))
                {
                    con.Open();
                    if (csv_path.Contains("\\t"))
                    {
                        csv_path = csv_path.Replace("\\t", "\\\\t");
                    }
                    string query = $"LOAD DATA INFILE \"{csv_path}\" INTO TABLE excel_data_db.exceldata;";
                    MySqlCommand cmd = new MySqlCommand(query, con);
                    result = cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
                throw (ex);
            }

            return result;
        }


        /// <summary>
        /// The DeleteFile.
        /// </summary>
        /// <param name="FilePath">The FilePath<see cref="string"/>.</param>
        /// <returns>The <see cref="bool"/>.</returns>
        private static bool DeleteFile(string FilePath)
        {
            try
            {
                bool IsFind = File.Exists(FilePath);
                if (IsFind)
                {
                    File.Delete(FilePath);
                }
                else
                {
                    throw new IOException("指定的文件不存在");
                }
                return true;
            }
            catch (Exception exc)
            {
                throw new Exception(exc.Message);
            }
        }

        /// <summary>
        /// The QuertExcel.
        /// </summary>
        private static void QuertExcel()
        {
            Process[] excels = Process.GetProcessesByName("EXCEL");
            foreach (var item in excels)
            {
                item.Kill();
            }
        }
    }
}
