using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            Utility.Generate.ExcelName("Aravind");
            Utility.Generate.Excel("S.No", "Name","City","Status");
            Utility.Generate.Excel(1, "Aravind", "dfcs", "sf");
            Utility.Generate.Excel(2, "aa", "efs", "sf");
            Utility.Generate.Excel(3, "ss", "sdf", "esf","dcsdvsdddddddddddddddddddddd");
            Utility.Generate.SaveExcel("Aravind");
        }
    }
}
