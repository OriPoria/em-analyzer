using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using EM_Analyzer.ModelClasses;
using System.Windows.Forms;

namespace EM_Analyzer.Services
{
    class ExcelsService
    {
        public static void CreateExcelFromStringTable<T>(string fileName, List<T> table)//List<string[]> table)
        {
            using (var wb = new XLWorkbook())
            {
                var ws = wb.Worksheets.Add("Inserting Tables");
                ws.Cell(1, 1).InsertTable(table);
                var titlesStyle = wb.Style;
                titlesStyle.Font.Bold = true;
                titlesStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                titlesStyle.Fill.BackgroundColor = XLColor.LightGray;
                titlesStyle.Font.FontSize = 14;
                ws.Range(1, 1, 1, ws.LastColumnUsed().ColumnNumber()).AsRange().AddToNamed("Titles");
                wb.NamedRanges.NamedRange("Titles").Ranges.Style = titlesStyle;

                ws.Columns().AdjustToContents();
                // wb.SaveAs(Application.StartupPath + "/" + fileName);
                wb.SaveAs(FixationsService.outputPath + "/" + FixationsService.textFileName.Substring(0,FixationsService.textFileName.IndexOf('.')) +" - " + fileName);
            }
        }

        public static List<IEnumerable<T>> ReadExcelFile<T>(string fileName)
        {
            XLWorkbook wb = new XLWorkbook(fileName);
            var ws = wb.Worksheets.First();
            var firstRowUsed = ws.FirstRowUsed();
            var categoryRow = firstRowUsed.RowUsed();
            List<IEnumerable<T>> table = new List<IEnumerable<T>>();

            // Move to the next row (it now has the titles)
            //string[] columnsHeaders = categoryRow.Cells().Select(cell => cell.GetValue<string>()).ToArray();
            categoryRow = categoryRow.RowBelow();
            while (!categoryRow.IsEmpty())
            {
                // table.Add(categoryRow.Cells().Select(cell => cell.GetValue<T>()).ToArray());
                table.Add(categoryRow.Cells().Select(cell => cell.GetValue<T>()));
                categoryRow = categoryRow.RowBelow();
            }

            return table;
        }
    }
}
