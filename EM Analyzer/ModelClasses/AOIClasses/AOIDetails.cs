using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EM_Analyzer.Enums;
using EM_Analyzer.ExcelLogger;
using EM_Analyzer.ModelClasses;
using EM_Analyzer.Services;

namespace EM_Analyzer.ModelClasses.AOIClasses
{
    public class AOIDetails : IAOI
    {
        public static Dictionary<string, AOIDetails> nameToAOIPhrasesDetailsDictionary = new Dictionary<string, AOIDetails>();
        public static Dictionary<int, string> groupPhraseToSpecialName = new Dictionary<int, string>();
        
        public static Dictionary<string, AOIDetails> nameToAOIWordsDetailsDictionary = new Dictionary<string, AOIDetails>();
        public static Dictionary<int, string> groupWordToSpecialName = new Dictionary<int, string>();

        //        public static bool isAOIIncludeStimulus = false;
        public string Stimulus { get; set; }
        public int Name { get; set; }
        public int Group { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double H { get; set; }
        public double L { get; set; }
        public double AOI_Coverage_In_Percents { get; set; }
        public bool IsProper { get; set; }
        public string DictionaryKey { get; private set; }
        public double AOI_Size_X { get; set; }
        public string SpecialNmae { get; set; }

        public AOIDetails(IEnumerable<string> details, uint lineNumber, AOITypes type)
        {
            AOI_Coverage_In_Percents = -1;
            AOI_Size_X = -1;
            IsProper = true;
            IEnumerator<string> enumerator = details.GetEnumerator();
            //string dictionaryKey = "";

            enumerator.MoveNext();
            int count = details.Count();
            Stimulus = enumerator.Current;
            enumerator.MoveNext();
            DictionaryKey = Stimulus;
            
            try
            {
                Name = int.Parse(enumerator.Current);
            }
            catch
            {
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("Name", enumerator.Current, lineNumber));
                IsProper = false;
            }
            enumerator.MoveNext();

            try
            {
                Group = int.Parse(enumerator.Current);
            }
            catch
            {
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("Group", enumerator.Current, lineNumber));
                IsProper = false;
            }
            enumerator.MoveNext();

            try
            {
                X = double.Parse(enumerator.Current);
            }
            catch
            {
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("X", enumerator.Current, lineNumber));
                IsProper = false;
            }
            enumerator.MoveNext();

            try
            {
                Y = double.Parse(enumerator.Current);
            }
            catch
            {
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("Y", enumerator.Current, lineNumber));
                IsProper = false;
            }
            enumerator.MoveNext();

            try
            {
                H = double.Parse(enumerator.Current);
            }
            catch
            {
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("H", enumerator.Current, lineNumber));
                IsProper = false;
            }
            enumerator.MoveNext();

            try
            {
                L = double.Parse(enumerator.Current);
            }
            catch
            {
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("L", enumerator.Current, lineNumber));
                IsProper = false;
            }
            enumerator.MoveNext();
            
            // to add AOI with special name
            if (count >= 8 && enumerator.Current != null) 
            {
                SpecialNmae = enumerator.Current;
                if (type == AOITypes.Phrases)
                    groupPhraseToSpecialName[Group] = SpecialNmae;
                else if (type == AOITypes.Words)
                    groupWordToSpecialName[Group] = SpecialNmae;

            }

            if (IsProper)
            {
                DictionaryKey = Name + DictionaryKey;
                if (type == AOITypes.Phrases)
                {
                    AOIsService.nameToAOIPhrasesDictionary[DictionaryKey] = this;
                    nameToAOIPhrasesDetailsDictionary[DictionaryKey] = this;
                } else if (type == AOITypes.Words)
                {
                    AOIsService.nameToAOIWordsDictionary[DictionaryKey] = this;
                    nameToAOIWordsDetailsDictionary[DictionaryKey] = this;
                }
            }
            

        }

        public static void LoadAllAOIPhraseFromFile(string fileName)
        {
            List<IEnumerable<string>> table = ExcelsService.ReadExcelFile<string>(fileName);
//            isAOIIncludeStimulus = false;
            IEnumerable<string> first = table.FirstOrDefault();
//            if (first.Count() >= 7)
//                isAOIIncludeStimulus = true;
            uint lineNumber = 1;
            foreach (IEnumerable<string> details in table)
            {
                new AOIDetails(details, lineNumber, AOITypes.Phrases);
                lineNumber++;
            }
        }
        public static void LoadAllAOIWordFromFile(string fileName)
        {
            List<IEnumerable<string>> table = ExcelsService.ReadExcelFile<string>(fileName);
            //            isAOIIncludeStimulus = false;
            IEnumerable<string> first = table.FirstOrDefault();
            //            if (first.Count() >= 7)
            //                isAOIIncludeStimulus = true;
            uint lineNumber = 1;
            foreach (IEnumerable<string> details in table)
            {
                List<string> detailsStr = details.ToList();
                detailsStr[2] = detailsStr[1];

                new AOIDetails(detailsStr, lineNumber, AOITypes.Words);
                lineNumber++;
            }
        }


        public double DistanceToAOI(Fixation fixation)
        {
            // The details of the AOI position is the top left point, width and height
            double right_border, left_border, upper_border, buttom_border, fixation_X, fixation_Y;
            right_border = X + L;
            left_border = X;
            buttom_border = Y + H;
            upper_border = Y;
            fixation_X = fixation.Fixation_Position_X;
            fixation_Y = fixation.Fixation_Position_Y;

            if (fixation_X >= left_border && fixation_X <= right_border) //Inside X Bounds
            {
                if (fixation_Y >= upper_border && fixation_Y <= buttom_border)//Inside Y Bounds
                    return 0;
                else //Outside Y Bounds
                {
                    if (fixation_Y < upper_border)
                    {
                        return upper_border - fixation_Y;
                    }
                    else //fixation_Y > buttom_border
                    {
                        return fixation_Y - buttom_border;
                    }
                }
            }
            else //Outside X Bounds
            {
                if (fixation_Y >= upper_border && fixation_Y <= buttom_border)//Inside Y Bounds
                {
                    if (fixation_X < left_border)
                    {
                        return left_border - fixation_X;
                    }
                    else //fixation_X > right_border
                    {
                        return fixation_X - right_border;
                    }
                }
                else //Outside Y Bounds (And X Bounds)
                {
                    double Y_Distance, X_Distance;
                    if (fixation_Y < upper_border)
                    {
                        Y_Distance = upper_border - fixation_Y;
                    }
                    else //fixation_Y > buttom_border
                    {
                        Y_Distance = fixation_Y - buttom_border;
                    }

                    if (fixation_X < left_border)
                    {
                        X_Distance = left_border - fixation_X;
                    }
                    else //fixation_X > right_border
                    {
                        X_Distance = fixation_X - right_border;
                    }

                    return Math.Sqrt(Math.Pow(X_Distance, 2) + Math.Pow(Y_Distance, 2));
                }
            }
        }


        private static Log CreateLogForFieldValidation(string fieldName, string valueFound, uint lineNumber)
        {
            return new Log() { FileName = FixationsService.phrasesExcelFileName, LineNumber = lineNumber, Description = "The Value Of Field " + fieldName + " Is Not Valid!!!" + Environment.NewLine + "The Value Found Is: " + valueFound };
        }
    }
}
