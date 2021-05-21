using EM_Analyzer.Enums;
using EM_Analyzer.ExcelsFilesMakers;
using EM_Analyzer.ModelClasses;
using EM_Analyzer.ModelClasses.AOIClasses;
using EM_Analyzer.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace EM_Analyzer
{
    public class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            int chosenOption;
            string input;
            bool isOptionOK;
            do
            {
                Console.WriteLine("Choose An Option: ");
                Console.WriteLine("1. Do The Full Process (Get Both The AOI And The Fixations And Create All The Excel Files).");
                Console.WriteLine("2. Do The Process Without Filtering Exceptions.");
                Console.WriteLine("3. Only Filtering Exceptions.");
                input = Console.ReadLine();
                isOptionOK = int.TryParse(input, out chosenOption);
                if (!isOptionOK)
                    Console.WriteLine("Please Choose A Valid Option!!!");
            } while (!isOptionOK);

            
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

            if (chosenOption == 3)
            {
                FixationsService.textFileName = excelFilePath.Substring(excelFilePath.LastIndexOf(@"\") + 1);
                FixationsService.outputPath = excelFilePath.Substring(0, excelFilePath.LastIndexOf(@"\"));
                var collection = ExcelsService.GetObjectsFromExcel(excelFilePath);
                foreach (var aoi in collection)
                {
                    aoi.CreateAIOClassAfterCoverage();
                }
                ExcelsFilesMakers.SecondFileConsideringCoverage.MakeExcelFile();
                return;
            }
            Thread readingExcelFile;
            readingExcelFile = new Thread(() =>
            {
                AOIDetails.LoadAllAOIFromFile(excelFilePath);
            });
            
            readingExcelFile.Start();

            string specialTextFilePath = "";
            
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
                    specialTextFilePath = openFileDialog.FileName;
                }
                else
                {
                    return;
                }
            }
            string wordsTextFilePath = "";

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
                    wordsTextFilePath = openFileDialog.FileName;
                }
                else
                {
                    return;
                }
            }

            readingExcelFile.Join();
            FixationsService.excelFileName = excelFilePath.Substring(excelFilePath.LastIndexOf(@"\") + 1);
            FixationsService.textFileName = specialTextFilePath.Substring(specialTextFilePath.LastIndexOf(@"\") + 1);
            FixationsService.outputPath = excelFilePath.Substring(0, excelFilePath.LastIndexOf(@"\"));
            ReadTextFiles(specialTextFilePath, wordsTextFilePath);
            FixationsService.DealWithSeparatedAOIs();
            FixationsService.SortDictionary();
            if (int.Parse(ConfigurationService.RemoveFixationsAppearedBeforeFirstAOI) == 1)
                FixationsService.CleanAllFixationBeforeFirstAOI();
            FixationsService.SearchForExceptions();

            ExcelsFilesMakers.FirstFileAfterProccessing.MakeExcelFile();
            Console.WriteLine("First File: " + ConfigurationService.FirstExcelFileName + " Finished!!! ");
            FixationsService.DealWithExceptions();
            ExcelsFilesMakers.SecondFileAfterProccessing.MakeExcelFile();
            if (chosenOption == 1)
            {
                ExcelsFilesMakers.SecondFileConsideringCoverage.MakeExcelFile();
                Console.WriteLine("Second File: " + ConfigurationService.SecondExcelFileName + " Finished!!! ");
            }
            ExcelsFilesMakers.ThirdFileAfterProccessing.MakeExcelFile();
            Console.WriteLine("Third File: " + ConfigurationService.ThirdExcelFileName + " Finished!!! ");


        }

        private static void ReadTextFiles(string specialFilePath, string wordFilePath)
        {
            string[] lines = File.ReadAllLines(specialFilePath);
            lines = lines.Where(line => line.Trim().Count() > 0).ToArray();
            string[] wordsLines = File.ReadAllLines(wordFilePath);
            wordsLines = wordsLines.Where(line => line.Trim().Count() > 0).ToArray();
            for (int i = 1; i < wordsLines.Length; i++)
            {
                string[] columns = wordsLines[i].Split('\t');
                string aoiName = columns[9].Split(' ')[1];
                lines[i] = lines[i] + $"\t{aoiName}";
            }
            FixationsService.tableColumns = lines[0].Split('\t').Select(column=>column.Trim().ToLower()).ToList();
            FixationsService.InitializeColumnIndexes();
            //i begin from 2, the first line is categors
            for (uint j = 2; j < lines.Length; j++)
            {
                string[] currentRow = lines[j].Split('\t');
                try
                {
                    Fixation.CreateFixationFromStringArray(currentRow, j);
                }
                catch
                {
                    MessageBox.Show("There is a problem with the Text File In Line: " + j + "\n Content: \"" + lines[j] + "\"");
                }
            }
            var x = FixationsService.fixationSetToFixationListDictionary;
            // Deletes Double Fixations (With The Same Index).
            IEnumerable<string> participants = FixationsService.fixationSetToFixationListDictionary.Keys.ToList();
            foreach (string participant in participants)
            {
                FixationsService.fixationSetToFixationListDictionary[participant] = FixationsService.fixationSetToFixationListDictionary[participant].GroupBy(fix=>fix.Index).Select(g=>g.First()).ToList();
            }
        }

        private static void SortTableByIndex(List<string[]> table)
        {
            table.Sort((a, b) => a[TextFileColumnIndexes.Index].CompareTo(b[TextFileColumnIndexes.Index]));
        }

        private void FirstDataProccessing(List<string[]> fixationTable)
        {

        }
    }
}
