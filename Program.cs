using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;

namespace read_write_excel
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel excel = new Excel("", 1); //Insert path to the excel file

            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet1 = workbook.CreateSheet("Sheet1");

            string[,] polje = new string[10000, 7];

            IRow pomocniRow;
            int brojacRedova = 1;
            for (int i = 1; i < 581; i++)
            {
                
                string id = excel.ReadCell(i, 1);
                string type = excel.ReadCell(i, 2);
                string title = excel.ReadCell(i, 3);
                string authors = excel.ReadCell(i, 4);
                string aff = excel.ReadCell(i, 5);
                string emails = excel.ReadCell(i, 6);

                string[] poljeAutora = authors.Split('|');
                string[] poljeAff = aff.Split('|');
                string[] poljeEmails = emails.Split('|');


                for (int z = 0; z<poljeAutora.Count(); z++)
                {
                    pomocniRow = sheet1.CreateRow(brojacRedova);
                    Console.WriteLine(poljeAutora.Count());
                    pomocniRow.CreateCell(1).SetCellValue(id);
                    pomocniRow.CreateCell(2).SetCellValue(type);
                    pomocniRow.CreateCell(3).SetCellValue(title);
                    pomocniRow.CreateCell(4).SetCellValue(poljeAutora[z]);
                    pomocniRow.CreateCell(5).SetCellValue(poljeAff[z]);
                    pomocniRow.CreateCell(6).SetCellValue(poljeEmails[z]);
                    brojacRedova++;
                }
            }

            FileStream sw = File.Create(""); //Insert name of the excel file that will be created - dont forget .xlsx
            workbook.Write(sw);
            sw.Close();

            Console.ReadKey();
        }
    }
}
