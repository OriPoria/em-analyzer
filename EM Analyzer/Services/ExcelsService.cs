using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using EM_Analyzer.ModelClasses;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using OfficeOpenXml.Style;
using System;
using System.Reflection;
using System.Windows.Forms;

namespace EM_Analyzer.Services
{

    public class EpplusIgnore : Attribute { }
    public class XLColumn : Attribute
    {
        public string Header { get; }

        public XLColumn(string Header)
        {
            this.Header = Header;
        }
    }

    public static class Extensions
    {
        public static ExcelRangeBase LoadFromCollectionFiltered<T>(this ExcelRangeBase @this, IEnumerable<T> collection)//, bool PrintHeaders, TableStyles styles) where T : class
        {
            MemberInfo[] membersToInclude = typeof(T)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(p => !Attribute.IsDefined(p, typeof(EpplusIgnore)))
                .ToArray();

            return @this.LoadFromCollection(collection, true,
                TableStyles.Light1, BindingFlags.Instance | BindingFlags.Public,
                membersToInclude);
        }

    }

    public class ExcelsService
    {
        public static void CreateExcelFromStringTable<T>(string fileName, List<T> table)//List<string[]> table)
        {
            using (var wb = new ExcelPackage())
            {
                //var ws = wb.Worksheets.Add("Inserting Tables");
                ExcelWorksheet ws = wb.Workbook.Worksheets.Add("Inserting Tables");
                ExcelRangeBase range = ws.Cells[1, 1].LoadFromCollectionFiltered(table);//,true,TableStyles.Medium1);
                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                ws.Cells[ws.Dimension.Address].Style.WrapText = true;
                ws.Cells[ws.Dimension.Address].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Cells[ws.Dimension.Address].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                DialogResult dialogResult; // = DialogResult.Retry;
                do
                {
                    try
                    {
                        wb.SaveAs(new FileInfo(FixationsService.outputPath + "/" + FixationsService.textFileName.Substring(0, FixationsService.textFileName.IndexOf('.')) + " - " + fileName + ConfigurationService.ExcelFilesExtension));
                        dialogResult = DialogResult.Abort;
                    }
                    catch (InvalidOperationException e)
                    {
                        Console.Write(e.InnerException.InnerException.Message);
                        string errorDescription = "";
                        errorDescription += e.InnerException?.InnerException?.Message + Environment.NewLine;
                        errorDescription += "Check If The File We Trying to overwrite is already open!!!";
                        dialogResult = MessageBox.Show(errorDescription, "Error In Saving File " + fileName, MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                    }
                } while (dialogResult == DialogResult.Retry);
            }
        }

        public static List<IEnumerable<T>> ReadExcelFile<T>(string fileName)
        {
            List<IEnumerable<T>> table = new List<IEnumerable<T>>();
            using (var wb = new ExcelPackage(new FileInfo(fileName)))
            {
                ExcelWorksheet ws = wb.Workbook.Worksheets.First();
                int firstRowUsed = ws.Dimension.Start.Row;
                int lastColUsed = ws.Dimension.End.Column;
                ExcelRow categoryRow = ws.Row(firstRowUsed);

                // Move to the next row (it now has the titles)
                for (int currentRow = firstRowUsed + 1 ; currentRow < ws.Dimension.End.Row ; currentRow++)
                {
                    ExcelRow row = ws.Row(currentRow);
                    ExcelRange range = ws.Cells[currentRow, 1, currentRow, lastColUsed];
                    table.Add(range.Select(cell => cell.GetValue<T>()).ToList());
                }
            }
            //XLWorkbook wb = new XLWorkbook(fileName);
            //var ws = wb.Worksheets.First();
            //var firstRowUsed = ws.FirstRowUsed();
            //var categoryRow = firstRowUsed.RowUsed();
            //List<IEnumerable<T>> table = new List<IEnumerable<T>>();

            //// Move to the next row (it now has the titles)
            ////string[] columnsHeaders = categoryRow.Cells().Select(cell => cell.GetValue<string>()).ToArray();
            //categoryRow = categoryRow.RowBelow();
            //while (!categoryRow.IsEmpty())
            //{
            //    // table.Add(categoryRow.Cells().Select(cell => cell.GetValue<T>()).ToArray());
            //    table.Add(categoryRow.Cells().Select(cell => cell.GetValue<T>()));
            //    categoryRow = categoryRow.RowBelow();
            //}

            return table;
            //return null;
        }
    }
}
