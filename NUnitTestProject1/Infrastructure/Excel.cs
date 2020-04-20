using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using ClosedXML.Excel;

namespace NUnitTestProject1
{
    class Excel
    {
        public static void SaveToExcel(IEnumerable<All> all)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("All");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "Получател";
                worksheet.Cell(currentRow, 2).Value = "Подател";
                

                foreach (var item in all)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = item.Sender;
                    worksheet.Cell(currentRow, 2).Value = item.Reciever;
                    
                }

                var fileName = Path.Combine(Environment.CurrentDirectory, "Data\\speedy.xlsx");
                workbook.SaveAs(fileName);

                //using (var stream = new MemoryStream())
                //{
                //workbook.SaveAs( (stream);
                //var content = stream.ToArray();
                //stream.Write

                //return File(
                //    content,
                //    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                //    "all-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".xlsx");
                //}
            }
        }

        public static void ReadFromExcelFile()
        {
            var fileName = Path.Combine(Environment.CurrentDirectory, "Data\\speedy.xlsx");
            //var fileName = Path.Combine(@"C:\Users\smasl\source\repos\ReadFromExcel\ReadFromExcel\Data", "\\speedy.xlsx");
            //C:\Users\smasl\source\repos\ReadFromExcel\ReadFromExcel\Data
            var workbook = new XLWorkbook(fileName);
            var ws1 = workbook.Worksheet(1);
            int iRow = 1;
            while (!ws1.Cell(iRow, 1).IsEmpty())
            {
                var row = "";
                int iColumn = 1;
                while (!ws1.Cell(iRow, iColumn).IsEmpty())
                {
                    row = row + ws1.Cell(iRow, iColumn).Value.ToString() + ",";
                    //string tempValue = ws1.Cell(iRow, iColumn).Value.ToString();
                    //Console.OutputEncoding = Encoding.UTF8;
                    //Console.WriteLine(ws1.Cell(iRow, iColumn).Value);
                    iColumn++;
                }
                //Console.WriteLine(row);
                iRow++;
            }
        }
    }
}
