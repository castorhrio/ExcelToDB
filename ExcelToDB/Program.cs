namespace ExcelToDB
{
    using System;

    /// <summary>
    /// Defines the <see cref="Program" />.
    /// </summary>
    internal class Program
    {
        /// <summary>
        /// The Main.
        /// </summary>
        /// <param name="args">The args<see cref="string[]"/>.</param>
        internal static void Main(string[] args)
        {
            string file_path = @"D:\excel_test\test.xls";
            string result = ExcelHelper.XLSSavesaCSV(file_path);
            //var dt_result = ExcelHelper.OpenCSV(result);

            var newCon = "server=127.0.0.1; port=3306; database=excel_data_db;uid=root;pwd=123;";
            int result_count = ExcelHelper.SqlBulkCopyInsert(newCon, result);
            Console.WriteLine(result_count);
            Console.ReadLine();
        }
    }
}
