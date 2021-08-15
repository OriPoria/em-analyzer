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
using static EM_Analyzer.ExcelsFilesMakers.SecondFileAfterProccessing;
using static EM_Analyzer.ExcelsFilesMakers.SecondFileConsideringCoverage;

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

            
            string phrasesExcelFilePath = "";
            string wordsExcelFilePath = "";

            while (phrasesExcelFilePath == "" || wordsExcelFilePath == "")
            {
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
                        if (openFileDialog.FileName.EndsWith("_c.xlsx"))
                            phrasesExcelFilePath = openFileDialog.FileName;
                        else if (openFileDialog.FileName.EndsWith("_w.xlsx"))
                            wordsExcelFilePath = openFileDialog.FileName;
                    }
                    else
                    {
                        return;
                    }

                }

            }


            if (chosenOption == 3)
            {
                FixationsService.phrasesTextFileName = phrasesExcelFilePath.Substring(phrasesExcelFilePath.LastIndexOf(@"\") + 1);
                FixationsService.wordsTextFileName = wordsExcelFilePath.Substring(phrasesExcelFilePath.LastIndexOf(@"\") + 1);
                FixationsService.outputPath = phrasesExcelFilePath.Substring(0, phrasesExcelFilePath.LastIndexOf(@"\"));
                string[] excelPathes = { phrasesExcelFilePath, wordsExcelFilePath };
                int i = 0;
                foreach (AOITypes type in (AOITypes[])Enum.GetValues(typeof(AOITypes)))
                {
                    var collection = ExcelsService.GetObjectsFromExcel(excelPathes[i]);
                    foreach (var aoi in collection)
                    {
                        aoi.CreateAIOClassAfterCoverage();
                    }
                    SecondFileConsideringCoverage.currentType = type;
                    ExcelsFilesMakers.SecondFileConsideringCoverage.MakeExcelFile();
                    i++;
                }

                return;
            }
            Thread readingExcelFile;
            readingExcelFile = new Thread(() =>
            {
                AOIDetails.LoadAllAOIPhraseFromFile(phrasesExcelFilePath);
                AOIDetails.LoadAllAOIWordFromFile(wordsExcelFilePath);
            });


            readingExcelFile.Start();

            string phrasesTextFilePath = "";
            string wordsTextFilePath = "";
            while (phrasesTextFilePath == "" || wordsTextFilePath == "")
            {
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
                        if (openFileDialog.FileName.EndsWith("_c.txt"))
                            phrasesTextFilePath = openFileDialog.FileName;
                        else if (openFileDialog.FileName.EndsWith("_w.txt"))
                            wordsTextFilePath = openFileDialog.FileName;
                    }
                    else
                    {
                        return;
                    }
                }
            }


            readingExcelFile.Join();
            FixationsService.phrasesExcelFileName = phrasesExcelFilePath.Substring(phrasesExcelFilePath.LastIndexOf(@"\") + 1);
            FixationsService.wordsExcelFileName = wordsExcelFilePath.Substring(wordsExcelFilePath.LastIndexOf(@"\") + 1);
            FixationsService.phrasesTextFileName = phrasesTextFilePath.Substring(phrasesTextFilePath.LastIndexOf(@"\") + 1);
            FixationsService.wordsTextFileName = wordsTextFilePath.Substring(wordsTextFilePath.LastIndexOf(@"\") + 1);
            FixationsService.outputPath = phrasesExcelFilePath.Substring(0, phrasesExcelFilePath.LastIndexOf(@"\"));
            ReadTextFiles(phrasesTextFilePath, wordsTextFilePath);
            FixationsService.DealWithSeparatedAOIs();
            FixationsService.SortDictionary();
            FixationsService.SortWordIndexDictionary();
            FixationsService.UnifyDictionaryWithWordIndex();
            if (int.Parse(ConfigurationService.RemoveFixationsAppearedBeforeFirstAOI) == 1)
                FixationsService.CleanAllFixationBeforeFirstAOI();
            FixationsService.SearchForExceptions();

            ExcelsFilesMakers.FirstFileAfterProccessing.MakeExcelFile();
            Console.WriteLine("First File: " + ConfigurationService.FirstExcelFileName + " Finished!!! ");
            FixationsService.DealWithExceptions();
            foreach (AOITypes type in (AOITypes[]) Enum.GetValues(typeof(AOITypes)))
            {
                SecondFileAfterProccessing.currentType = type;
                ExcelsFilesMakers.SecondFileAfterProccessing.MakeExcelFile();
                     if (chosenOption == 1)
                     {
                         SecondFileConsideringCoverage.currentType = type;
                         ExcelsFilesMakers.SecondFileConsideringCoverage.MakeExcelFile();
                         AIOClassAfterCoverage.allInstances.Clear();

                    }
                AOIClass.instancesDictionary.Clear();

            }
            Console.WriteLine("Second File: " + ConfigurationService.SecondExcelFileName + " Finished!!! ");

            ExcelsFilesMakers.ThirdFileAfterProccessing.MakeExcelFile();
            Console.WriteLine("Third File: " + ConfigurationService.ThirdExcelFileName + " Finished!!! ");


        }

        private static void ReadTextFiles(string phraseFilePath, string wordFilePath)
        {
            string[] lines = File.ReadAllLines(phraseFilePath);

            lines = lines.Where(line => line.Trim().Count() > 0).ToArray();

            FixationsService.tableColumns = lines[0].Split('\t').Select(column=>column.Trim().ToLower()).ToList();
            FixationsService.InitializeColumnIndexes();
            //i begin from 1, the first line is categors
            for (uint j = 1; j < lines.Length; j++)
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
            // Deletes Double Fixations (With The Same Index).
            IEnumerable<string> participants = FixationsService.fixationSetToFixationListDictionary.Keys.ToList();
            foreach (string participant in participants)
            {
                FixationsService.fixationSetToFixationListDictionary[participant] = FixationsService.fixationSetToFixationListDictionary[participant].GroupBy(fix=>fix.Index).Select(g=>g.First()).ToList();
            }
            string[] wordsLines = File.ReadAllLines(wordFilePath);
            wordsLines = wordsLines.Where(line => line.Trim().Count() > 0).ToArray();
            for (uint j = 1; j < wordsLines.Length; j++)
            {
                string[] columns = wordsLines[j].Split('\t');
                Fixation.CreateWordIndexFromStringArray(columns, j);
            }
            foreach (string participant in participants)
            {
                FixationsService.wordIndexSetToFixationListDictionary[participant] = FixationsService.wordIndexSetToFixationListDictionary[participant].GroupBy(fix => fix.Index).Select(g => g.First()).ToList();
            }
            Dictionary<string, List<Fixation>> fixatioDicTest = FixationsService.fixationSetToFixationListDictionary;
            Dictionary<string, List<WordIndex>> wordsDicTest = FixationsService.wordIndexSetToFixationListDictionary;
        }

        private static void SortTableByIndex(List<string[]> table)
        {
            table.Sort((a, b) => a[TextFileColumnIndexes.Index].CompareTo(b[TextFileColumnIndexes.Index]));
        }

    }
}
