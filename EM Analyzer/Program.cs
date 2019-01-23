using ForBarIlanResearch.Enums;
using ForBarIlanResearch.ModelClasses;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ForBarIlanResearch.Services;
using Excel = Microsoft.Office.Interop.Excel;

namespace ForBarIlanResearch
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string excelFilePath = "";

            using (OpenFileDialog openFileDialog = new OpenFileDialog
            {
                RestoreDirectory = true,
                Title = "Open The Excel File: ",
                Filter = "Excel files (*.xls*)|*.xls*",
                Multiselect = false,
                CheckFileExists = true,
                CheckPathExists = true
            })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    excelFilePath = openFileDialog.FileName;
                }
                else
                {
                    return;
                }
            }

            Thread readingExcelFile = new Thread(() => { AOIDetails.LoadAllAOIFromFile(excelFilePath); });
            readingExcelFile.Start();

            string textFilePath = "";

            using (OpenFileDialog openFileDialog = new OpenFileDialog
            {
                RestoreDirectory = true,
                Title = "Open The Text File: ",
                Filter = "Text files (*.txt)|*.txt",
                Multiselect = false,
                CheckFileExists = true,
                CheckPathExists = true
            })
            {
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    textFilePath = openFileDialog.FileName;
                }
                else
                {
                    return;
                }
            }

            readingExcelFile.Join();
            FixationsService.textName = excelFilePath.Substring(excelFilePath.LastIndexOf(@"\") + 1);
            FixationsService.outputPath = excelFilePath.Substring(0, excelFilePath.LastIndexOf(@"\"));
            readTextFile(textFilePath);
            FixationsService.SortDictionary();
            FixationsService.CleanAllFixationBeforeFirstAOI();
            FixationsService.SearchForExceptions();

            ExcelsFilesMakers.FirstFileAfterProccessing.makeExcelFile();
            Console.WriteLine("First File: " + ConfigurationService.getValue(ConfigurationService.First_Excel_File_Name) + " Finished!!! ");
            ExcelsFilesMakers.SecondFileAfterProccessing.makeExcelFile();
            Console.WriteLine("Second File: " + ConfigurationService.getValue(ConfigurationService.Second_Excel_File_Name) + " Finished!!! ");
            ExcelsFilesMakers.ThirdFileAfterProccessing.makeExcelFile();
            Console.WriteLine("Third File: " + ConfigurationService.getValue(ConfigurationService.Third_Excel_File_Name) + " Finished!!! ");


        }

        private static void readTextFile(string filePath)
        {
            string[] lines = File.ReadAllLines(filePath).Where(line => line.Trim().Count() > 0).ToArray();

            FixationsService.tableColumns = lines[0].Split('\t').ToList();

            for (int i = 1; i < lines.Length; i++)
            {
                string[] currentRow = lines[i].Split('\t');
                try
                {
                    Fixation.CreateFixationFromStringArray(currentRow);
                }
                catch
                {
                    MessageBox.Show("There is a problem with the Text File In Line: " + i + "\n Content: \"" + lines[i] + "\"");
                }
            }
        }

        private static void sortTableByIndex(List<string[]> table)
        {
            table.Sort((a, b) => a[(int)TableColumnsEnum.Index].CompareTo(b[(int)TableColumnsEnum.Index]));
        }

        private static void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void FirstDataProccessing(List<string[]> fixationTable)
        {

        }
    }
}
