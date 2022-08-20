using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using BetterConsoleTables;

namespace ConsoleToExcel.Utility
{
   public class Generate
    {
        public static XLWorkbook workbook = new XLWorkbook();
        public static ClosedXML.Excel.IXLWorksheet worksheet;
        //public static Table table = new Table("one", "two", "three");

        public static int ROW_COUNT = 1;

        public static void ExcelName(string SheetName)
        {
            worksheet = workbook.Worksheets.Add(SheetName);
        }
        public static void SaveExcel(string ExcelName )
        {
            workbook.SaveAs(ExcelName+".xlsx");
        }
        public static void Excel(params dynamic[] Data)
        {

            int i = 0;
            foreach (var item in Data)
            {
                var CELLDATA = GCN(i) + ROW_COUNT;
                worksheet.Cell(CELLDATA).Value = item;
                i++;
            }
            ROW_COUNT++;
            TABLE(Data);
        }

        public static void TABLE(params dynamic[] Data)
        {
            string val = "";
            foreach (var item in Data)
            {
                val = val + " " + item;
            }
            Console.WriteLine(val);
        }

        static string GCN(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];

            return value;
        }
    }
}
