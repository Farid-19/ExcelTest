using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ExcelTestThingie.Models;
using OfficeOpenXml;

namespace ExcelTestThingie.Service
{
    public class ExcelService
    {
        public ExcelService()
        {
        }


        public void AppendDuelistsToExcel(List<Duelist> duelists, string title)
        {
            var fileInfo = new FileInfo(@"Template\Duelists_Template.xlsx");

            // Open the workbook (or create it if it doesn't exist)
            using (var excelPackage = new ExcelPackage(fileInfo))
            {
                // Open up the worksheet.
                var workSheet = excelPackage.Workbook.Worksheets["Duelists"];

                // Title
                workSheet.Cells[1, 1].Value = title;

                // DateTime
                workSheet.Cells[2, 3].Value = DateTime.Now.ToString("d/M/yyyy HH:mm:ss");

                var row = 5;
                foreach (var duelist in duelists)
                {
                    workSheet.Cells[row, 1].Value = duelist.Country;
                    workSheet.Cells[row, 2].Value = duelist.Firstname;
                    workSheet.Cells[row, 3].Value = duelist.Lastname;
                    workSheet.Cells[row, 4].Value = duelist.Rank.ToString();
                    row++;
                }


                var excelName = $"D:\\{DateTime.Now:yy-MM-dd}_Duelists.xlsx";
                //Save and close the package.
                excelPackage.SaveAs(new FileInfo(excelName));
            }

        }
    }
}
