using System.Collections.Generic;
using System.IO;
using System.Linq;
using EM_Analyzer.ModelClasses;
using OfficeOpenXml;
using OfficeOpenXml.Table;
using System;
using System.Reflection;
using System.Windows.Forms;
using System.Diagnostics;
using System.ComponentModel;
using System.Threading;
using EM_Analyzer.ExcelsFilesMakers;
using System.Text.RegularExpressions;

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
                TableStyles.None,
                BindingFlags.Instance | BindingFlags.Public,
                membersToInclude);
        }

    }

    public class ExcelsService
    {
        // params:
        // file name - The type of the excel file, come from the configuration xml file parameter.
        public static void CreateExcelFromStringTable<T>(string fileName, IEnumerable<T> table, Func<ExcelWorksheet, int> editExcelFunc)
        {
            using (var wb = new ExcelPackage())
            {
                ExcelWorksheet ws = wb.Workbook.Worksheets.Add("Inserting Tables");
                string ss = FixationsService.phrasesTextFileName;
                string textDataName = FixationsService.phrasesTextFileName.Substring(0, FixationsService.phrasesTextFileName.Length - 6);
                String islogs = "Logs";
                String isFiltered = "AOI - Filtered";
                if (!fileName.Contains(islogs))
                {
                    ws.View.FreezePanes(2, 4);
                }  
                if (fileName.Contains(isFiltered))
                {
                    ws.View.FreezePanes(2, 6);
                }
                ExcelRangeBase range = ws.Cells[1, 1].LoadFromCollectionFiltered(table);

                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                editExcelFunc?.Invoke(ws);
                DialogResult dialogResult; // = DialogResult.Retry;
                do
                {
                    try
                    {
                        if (fileName.Contains(isFiltered))
                        {
                            string path = FixationsService.outputPath + "/" + textDataName + " - Filters";
                            if (!Directory.Exists(path))
                            {
                                DirectoryInfo di = Directory.CreateDirectory(path);
                            }
                            wb.SaveAs(new FileInfo(path + "/" + fileName + ConfigurationService.ExcelFilesExtension));
                        }
                        else
                            wb.SaveAs(new FileInfo(FixationsService.outputPath + "/" + textDataName + " - " + fileName + ConfigurationService.ExcelFilesExtension));
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
                for (int currentRow = firstRowUsed + 1 ; currentRow <= ws.Dimension.End.Row ; currentRow++)
                {
                    ExcelRow row = ws.Row(currentRow);
                    ExcelRange range = ws.Cells[currentRow, 1, currentRow, lastColUsed];
                    table.Add(range.Select(cell => cell.GetValue<T>()).ToList());
                }
            }

            return table;
        }

        public static List<AIOClassFromExcel> GetObjectsFromExcel(string excelFileName)
        {
            List<AIOClassFromExcel> list = new List<AIOClassFromExcel>();
            if (excelFileName != null)
            {
                try
                {
                    var fi = new FileInfo(excelFileName);
                    using (ExcelPackage package = new ExcelPackage(fi))
                    {
                        ExcelWorkbook workbook = package.Workbook;
                        if (workbook != null)
                        {
                            ExcelWorksheet worksheet = workbook.Worksheets.FirstOrDefault();
                            if (worksheet != null)
                            {
                                list = ImportExcelToList<AIOClassFromExcel>(worksheet);
                            }
                        }
                    }
                }
                catch (Exception err)
                {
                    //save error log
                }
            }
            return list;
        }

        public static List<T> ImportExcelToList<T>(ExcelWorksheet worksheet) where T : new()
        {
            //DateTime Conversion
            Func<double, DateTime> convertDateTime = new Func<double, DateTime>(excelDate =>
            {
                if (excelDate < 1)
                {
                    throw new ArgumentException("Excel dates cannot be smaller than 0.");
                }

                DateTime dateOfReference = new DateTime(1900, 1, 1);

                if (excelDate > 60d)
                {
                    excelDate = excelDate - 2;
                }
                else
                {
                    excelDate = excelDate - 1;
                }

                return dateOfReference.AddDays(excelDate);
            });

            ExcelTable table = null;

            if (worksheet.Tables.Any())
            {
                table = worksheet.Tables.FirstOrDefault();
            }
            else
            {
                //table = worksheet.Tables.Add(worksheet.Dimension, "tbl" + Guid.NewGuid().ToString());
                table = worksheet.Tables.Add(worksheet.Dimension, "tbl");

                ExcelAddressBase newaddy = new ExcelAddressBase(table.Address.Start.Row, table.Address.Start.Column, table.Address.End.Row + 1, table.Address.End.Column);

                //Edit the raw XML by searching for all references to the old address
                table.TableXml.InnerXml = table.TableXml.InnerXml.Replace(table.Address.ToString(), newaddy.ToString());
            }

            //Get the cells based on the table address
            List<IGrouping<int, ExcelRangeBase>> groups = table.WorkSheet.Cells[table.Address.Start.Row, table.Address.Start.Column, table.Address.End.Row, table.Address.End.Column]
                .GroupBy(cell => cell.Start.Row)
                .ToList();

            //Assume the second row represents column data types (big assumption!)
            List<Type> types = groups.Skip(1).FirstOrDefault().Select(rcell => rcell.Value.GetType()).ToList();

            //Get the properties of T
            List<PropertyInfo> modelProperties = new T().GetType().GetProperties().ToList();

            var colnames = groups.FirstOrDefault()
                .Select((hcell, idx) => new
                {
                    Name = hcell.Value.ToString().Replace("Pupil Diameter [mm]","Pupil_Diameter")
                                                 .Replace("AOI Size X [mm]","Mean_AOI_Size")
                                                 .Replace("AOI Coverage [%]","Mean_AOI_Coverage").
                                                  Replace(" ", "_").Replace("-","_"),
                    //Name = hcell.Value.ToString(),
                    index = idx
                }).ToList();

            //Everything after the header is data
            List<List<object>> rowvalues = groups
                .Skip(1) //Exclude header
                .Select(cg => cg.Select(c => c.Value).ToList()).ToList();

            
            //Create the collection container
            List<T> collection = new List<T>();
            foreach (List<object> row in rowvalues)
            {
                T tnew = new T();
                foreach (var colname in colnames)
                {
                    //This is the real wrinkle to using reflection - Excel stores all numbers as double including int
                    object val = row[colname.index];
                    Type type = types[colname.index];
                    PropertyInfo prop = modelProperties.FirstOrDefault(p => p.Name == colname.Name);

                    //If it is numeric it is a double since that is how excel stores all numbers
                    if (type == typeof(double))
                    {
                        //Unbox it
                        double unboxedVal = (double)val;

                        //FAR FROM A COMPLETE LIST!!!
                        if (prop.PropertyType == typeof(int))
                        {
                            prop.SetValue(tnew, (int)unboxedVal);
                        }
                        else if (prop.PropertyType == typeof(double))
                        {
                            prop.SetValue(tnew, unboxedVal);
                        }
                        else if (prop.PropertyType == typeof(DateTime))
                        {
                            prop.SetValue(tnew, convertDateTime(unboxedVal));
                        }
                        else if (prop.PropertyType == typeof(string))
                        {
                            prop.SetValue(tnew, val.ToString());
                        }
                        else
                        {
                            throw new NotImplementedException(string.Format("Type '{0}' not implemented yet!", prop.PropertyType.Name));
                        }
                    }
                    else
                    {
                        //Its a string
                        prop.SetValue(tnew, val);
                    }
                }
                collection.Add(tnew);
            }

            return collection;
        }
    }
}
