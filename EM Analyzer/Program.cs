using EM_Analyzer.Enums;
using EM_Analyzer.ModelClasses;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace EM_Analyzer
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
            FixationsService.excelFileName = excelFilePath.Substring(excelFilePath.LastIndexOf(@"\") + 1);
            FixationsService.textFileName = textFilePath.Substring(textFilePath.LastIndexOf(@"\") + 1);
            FixationsService.outputPath = excelFilePath.Substring(0, excelFilePath.LastIndexOf(@"\"));
            readTextFile(textFilePath);
            FixationsService.SortDictionary();
            if(int.Parse(ConfigurationService.RemoveFixationsAppearedBeforeFirstAOI)==1)
                FixationsService.CleanAllFixationBeforeFirstAOI();
            FixationsService.SearchForExceptions();
            FixationsService.DealWithExceptions();

            ExcelsFilesMakers.FirstFileAfterProccessing.makeExcelFile();
            Console.WriteLine("First File: " + ConfigurationService.FirstExcelFileName + " Finished!!! ");
            ExcelsFilesMakers.SecondFileAfterProccessing.makeExcelFile();
            Console.WriteLine("Second File: " + ConfigurationService.SecondExcelFileName + " Finished!!! ");
            ExcelsFilesMakers.ThirdFileAfterProccessing.makeExcelFile();
            Console.WriteLine("Third File: " + ConfigurationService.ThirdExcelFileName + " Finished!!! ");


        }

        private static void readTextFile(string filePath)
        {
            string[] lines = File.ReadAllLines(filePath).Where(line => line.Trim().Count() > 0).ToArray();

            FixationsService.tableColumns = lines[0].Split('\t').ToList();

            for (uint i = 1; i < lines.Length; i++)
            {
                string[] currentRow = lines[i].Split('\t');
                try
                {
                    Fixation.CreateFixationFromStringArray(currentRow, i);
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
