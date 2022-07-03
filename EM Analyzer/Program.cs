using OfficeOpenXml;
using OfficeOpenXml.Table;
using EM_Analyzer.Enums;
using EM_Analyzer.ExcelsFilesMakers;
using EM_Analyzer.ExcelsFilesMakers.ThirdFourFilter;
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
            ExcelWorkbook ws2 = new ExcelPackage(new FileInfo(@"C:\Users\oripo\Downloads\AOI_boundaries-ELIZABETH_1_L_QA+E_P1_1_TEXT_w.xlsx")).Workbook;
            int chosenOption = 1;
            string input;
            bool isOptionOK;
            bool testMode = false;
            if (!testMode)
            {
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


            }
            char createPerParticipantFilterFiles = 'n';
            isOptionOK = false;
            /*
            do
            {
                Console.WriteLine("Create output files per participant filer: (Y/N)");
                input = Console.ReadLine();
                createPerParticipantFilterFiles = input[0];
                if (createPerParticipantFilterFiles != 'Y' && createPerParticipantFilterFiles != 'N' && createPerParticipantFilterFiles != 'y' && createPerParticipantFilterFiles != 'n')
                    Console.WriteLine("Please Choose A Valid Option!!!");
                else
                    isOptionOK = true;
            } while (!isOptionOK);
            */
            string phrasesExcelFilePath = "";
            string wordsExcelFilePath = "";
            
            if (!testMode)
            {
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
                            else
                                Console.WriteLine("Please upload excel file that ends with _c or _w");
                        }
                        else
                        {
                            return;
                        }

                    }

                }
            }
            // Test mode- choose ahead the files
            else
            {
                phrasesExcelFilePath = @"C:\Users\oripo\Desktop\AOI_EM_c.xlsx";
                wordsExcelFilePath = @"C:\Users\oripo\Desktop\AOI_EM_w.xlsx";
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
                    SecondFileConsideringCoverage.MakeExcelFile();
                    i++;
                }

                return;
            }
            Thread readingExcelFile;
            readingExcelFile = new Thread(() =>
            {
                AOIDetails.LoadAllAOIPhraseFromFile(phrasesExcelFilePath);
                AOIWordDetails.LoadAllAOIWordFromFile(wordsExcelFilePath);
            });


            readingExcelFile.Start();
            string phrasesTextFilePath = "";
            string wordsTextFilePath = "";
            if (!testMode)
            {
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
                            else
                                Console.WriteLine("Please upload text file that ends with _c or _w");
                        }
                        else
                        {
                            return;
                        }
                    }
                }
            }
            // Test mode- choose ahead the files
            else
            {
    //            phrasesTextFilePath = @"C: \Users\oripo\Desktop\work\EyeTracker\8_trial2_c.txt";
         //       wordsTextFilePath = @"C: \Users\oripo\Desktop\work\EyeTracker\8_trial2_w.txt";
          //     phrasesTextFilePath = @"C: \Users\oripo\Downloads\par25_c.txt";
            //    wordsTextFilePath = @"C: \Users\oripo\Downloads\par25_w.txt";
                phrasesTextFilePath = @"C: \Users\oripo\Downloads\EM_c.txt";
                wordsTextFilePath = @"C: \Users\oripo\Downloads\EN_w.txt";

                //                phrasesTextFilePath = @"C: \Users\oripo\Desktop\work\EyeTracker\small_20lines_c.txt";
                //            wordsTextFilePath = @"C: \Users\oripo\Desktop\work\EyeTracker\small_20lines_w.txt";
                //           phrasesTextFilePath = @"C: \Users\oripo\Desktop\work\EyeTracker\small_c.txt";
                //         wordsTextFilePath = @"C: \Users\oripo\Desktop\work\EyeTracker\small_w.txt";


            }
            readingExcelFile.Join();
            FixationsService.phrasesExcelFileName = phrasesExcelFilePath.Substring(phrasesExcelFilePath.LastIndexOf(@"\") + 1);
            FixationsService.wordsExcelFileName = wordsExcelFilePath.Substring(wordsExcelFilePath.LastIndexOf(@"\") + 1);
            FixationsService.phrasesTextFileName = phrasesTextFilePath.Substring(phrasesTextFilePath.LastIndexOf(@"\") + 1);
            FixationsService.wordsTextFileName = wordsTextFilePath.Substring(wordsTextFilePath.LastIndexOf(@"\") + 1);
            FixationsService.outputPath = phrasesExcelFilePath.Substring(0, phrasesExcelFilePath.LastIndexOf(@"\"));
            FixationsService.outputTextString = FixationsService.phrasesTextFileName.Substring(0, FixationsService.phrasesTextFileName.Length - 6);

            ReadTextFilePhrase(phrasesTextFilePath);
            ReadTextFileWord(wordsTextFilePath);

            FixationsService.DealWithSeparatedAOIs();
            FixationsService.SortDictionaryByFixationsIndex();
            FixationsService.SortWordIndexDictionaryByFixationsIndex();
            int status = FixationsService.UnifyDictionaryWithWordIndex();
            if (status == -1)
                return;
            FixationsService.SetTextIndex();
            FixationsService.SortDictionary();
            /*
             * After SortDictionary function, FixationsService.fixationSetToFixationListDictionary hold 
             * the fixations with all the details, include AOI's and word index
             */

            // if AnalyzeExtent is 1-> create only preview of fixations

            if (int.Parse(ConfigurationService.AnalyzeExtent) == 1)
            {
                FixationsService.outputTextString += " - Preview";
                FixationsService.CleanFixationsForPreview();
            }
            // if AnalyzeExtent is 2-> create full outputs
            else if (int.Parse(ConfigurationService.AnalyzeExtent) == 2) 
            {
                if (int.Parse(ConfigurationService.RemoveFixationsAppearedBeforeFirstAOI) == 2)
                    FixationsService.CleanAllFixationBeforeFirstAOIInFirstPage();
                else if (int.Parse(ConfigurationService.RemoveFixationsAppearedBeforeFirstAOI) == 3)
                    FixationsService.CleanAllFixationBeforeFirstAOIInPage();
            }
            FixationsService.SearchForExceptions();
            FirstFileAfterProccessing.MakeExcelFile();
            Console.WriteLine("First File: " + ConfigurationService.FirstExcelFileName + " Finished!!! ");
            FixationsService.DealWithExceptions();
            foreach (AOITypes type in (AOITypes[])Enum.GetValues(typeof(AOITypes)))
            {
                SecondFileAfterProccessing.currentType = type;
                SecondFileConsideringCoverage.currentType = type;
                SecondFileAfterProccessing.MakeExcelFile();
                if (chosenOption == 1)
                {
                    SecondFileConsideringCoverage.MakeExcelFile();
                    AIOClassAfterCoverage.allInstances.Clear();
                }
                AOIClass.instancesDictionary.Clear();
                AOIClass.maxConditions = 0;
            }
            Console.WriteLine("Second File: " + ConfigurationService.SecondExcelFileName + " Finished!!! ");

            /*
             * From ThirdFileAfterProccessing we remove all the fixations with AOI < 1 -> those without AOI for calculations
             */
            ThirdFileAfterProccessing.MakeExcelFile();
            Console.WriteLine("Third File: " + ConfigurationService.ThirdExcelFileName + " Finished!!! ");

            FourthFileAfterProccessing.MakeExcelFile();
            Console.WriteLine("Fourth File: " + ConfigurationService.FourthExcelFileName + " Finished!!! ");

        // filter per participant
            ThirdFourthFilter.CreateDatasetFilterTrialText();
            if (createPerParticipantFilterFiles == 'y' || createPerParticipantFilterFiles == 'Y')
                ThirdFourthFilter.CreateFilesForTest();
            ThirdFileConsideringCoverage.MakeExcelFile();
            FourthFileConsideringCoverage.MakeExcelFile();
            Console.WriteLine("Third and Fourth Filter File: " + ConfigurationService.ThirdExcelFileName + " Finished!!! ");

        }

        private static void ReadTextFilePhrase(string phraseFilePath)
        {
            string[] lines = File.ReadAllLines(phraseFilePath);

            lines = lines.Where(line => line.Trim().Count() > 0).ToArray();

            FixationsService.tableColumns = lines[0].Split('\t').Select(column=>column.Trim().ToLower()).ToList();
            FixationsService.InitializeColumnIndexes();
            // i begin from 1, the first line is categors
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
                if (j % 1000 == 0)
                {
                    Console.Write("\rCompleted process of {0} fixations out of {1} from closure file", j, lines.Length);
                }
            }
            Console.Write("\r");
            Console.Write(new string(' ', Console.BufferWidth));
            Console.SetCursorPosition(0, Console.CursorTop - 1);

            Console.WriteLine("\rCompleted process of fixations from closure file");
        }

        private static void ReadTextFileWord(string wordFilePath)
        {
            IEnumerable<string> participants = FixationsService.fixationSetToFixationListDictionary.Keys.ToList();

            string[] wordsLines = File.ReadAllLines(wordFilePath);
            wordsLines = wordsLines.Where(line => line.Trim().Count() > 0).ToArray();
            for (uint j = 1; j < wordsLines.Length; j++)
            {
                string[] columns = wordsLines[j].Split('\t');
                WordIndex.CreateWordIndexFromStringArray(columns, j);
                if (j % 1000 == 0)
                {
                    Console.Write("\rCompleted process of {0} fixations out of {1} from words file", j, wordsLines.Length);
                }

            }

            foreach (string participant in participants)
            {
                FixationsService.wordIndexSetToFixationListDictionary[participant] = FixationsService.wordIndexSetToFixationListDictionary[participant].GroupBy(fix => fix.Index).Select(g => g.First()).ToList();
            }

            Console.Write("\r");
            Console.Write(new string(' ', Console.BufferWidth));
            Console.SetCursorPosition(0, Console.CursorTop - 1);

            Console.WriteLine("\rCompleted process of fixations from words file");

        }


    }
}
