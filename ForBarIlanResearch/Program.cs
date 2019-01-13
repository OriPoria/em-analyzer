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
            DateTime before = DateTime.Now;
            //string folderPath;
            //System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(() =>
            //{
            //    using (FolderBrowserDialog saveFileDialog = new FolderBrowserDialog
            //    {
            //        ShowNewFolderButton = true,
            //        Description = "Choose The Folder To Save To:"
            //        //CheckPathExists = true
            //    })
            //    {
            //        if (saveFileDialog.ShowDialog() == DialogResult.OK)
            //        {
            //            folderPath = saveFileDialog.SelectedPath;

            //        }
            //    }
            //});
            string excelFilePath="";

            System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(() =>
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    RestoreDirectory = true,
                    Title = "Open The Excel File: ",
                    Filter = "Excel files (*.xls*)|*.xls*",
                    Multiselect = false,
                    CheckFileExists = true,
                    CheckPathExists=true
                })
                {
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        excelFilePath = openFileDialog.FileName;

                    }
                }
            });

            //AOIDetails.LoadAllAOIFromFile(Application.StartupPath + "/Hanaka.xlsx");

            Thread readingExcelFile = new Thread(() => { AOIDetails.LoadAllAOIFromFile(excelFilePath); });
            readingExcelFile.Start();

            string textFilePath = "";

            System.Windows.Threading.Dispatcher.CurrentDispatcher.Invoke(() =>
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    RestoreDirectory = true,
                    Title = "Open The Text File: ",
                    Filter = "Text files (*.txt)|*.txt",
                    Multiselect=false,
                    CheckFileExists=true,
                    CheckPathExists = true
                })
                {
                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        textFilePath = openFileDialog.FileName;

                    }
                }
            });

            

            readingExcelFile.Join();
            FixationsService.textName= excelFilePath.Substring(excelFilePath.LastIndexOf(@"\") + 1);
            FixationsService.outputPath = excelFilePath.Substring(0, excelFilePath.LastIndexOf(@"\"));
            readTextFile(textFilePath);
            FixationsService.SortDictionary();
            FixationsService.CleanAllFixationBeforeFirstAOI();
            FixationsService.SearchForExceptions();

            
            /*
            Thread makeFirstExcelFile = ExcelsFilesMakers.FirstFileAfterProccessing.makeExcelFile();
            FixationsService.DealWithExceptions();
            Task makeSecondExcelFile = new Task(() => ExcelsFilesMakers.SecondFileAfterProccessing.makeExcelFile());
            makeSecondExcelFile.Start();
            Task makeThirdExcelFile = new Task(() => ExcelsFilesMakers.ThirdFileAfterProccessing.makeExcelFile());
            makeThirdExcelFile.Start();

            makeFirstExcelFile.Join();
            makeSecondExcelFile.Wait();
            makeThirdExcelFile.Wait();
            */

            ExcelsFilesMakers.FirstFileAfterProccessing.makeExcelFile();
            Console.WriteLine("First File: " + ConfigurationService.getValue(ConfigurationService.First_Excel_File_Name)+ " Finished!!! ");
            ExcelsFilesMakers.SecondFileAfterProccessing.makeExcelFile();
            Console.WriteLine("Second File: " + ConfigurationService.getValue(ConfigurationService.Second_Excel_File_Name) + " Finished!!! ");
            ExcelsFilesMakers.ThirdFileAfterProccessing.makeExcelFile();
            Console.WriteLine("Third File: " + ConfigurationService.getValue(ConfigurationService.Third_Excel_File_Name) + " Finished!!! ");
            //Console.WriteLine((DateTime.Now - before).TotalMilliseconds);
            //Console.ReadKey();


        }

        private static void readTextFile(string filePath)
        {
            //DateTime before = DateTime.Now;
            //string fileName = "AOI Statistics - Single hanaka.txt";
            //string[] lines = File.ReadAllLines(Application.StartupPath + "/" + fileName);
            string[] lines = File.ReadAllLines(filePath);
            //int columns = lines[0].Count(c => (c == '\t')) + 1;
            FixationsService.tableColumns = lines[0].Split('\t').ToList();
            FixationsService.InitializeColumnIndexes();


            for (int i = 1; i < lines.Length; i++)
            {
                string[] currentRow = lines[i].Split('\t');
                Fixation.CreateFixationFromStringArray(currentRow);
                //currentFixation = Fixation.CreateFixationFromStringArray(currentRow);
                //fixations[i - 1] = currentFixation;
            }
            //fixations = fixations.Where((f) => f != null).ToArray();
            //Console.WriteLine((DateTime.Now - before).Milliseconds);
            //return null;
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
