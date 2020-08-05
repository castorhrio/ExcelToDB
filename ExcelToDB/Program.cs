using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToDB
{
    class Program
    {
        static void Main(string[] args)
        {
            string file_path = @"E:\test.xls";
            string result = ExcelHelper.XLSSavesaCSV(file_path);
            var dt_result = ExcelHelper.OpenCSV(result);
            ExcelHelper.ForDataTable(dt_result);
            Console.ReadLine();
        }
    }
}
