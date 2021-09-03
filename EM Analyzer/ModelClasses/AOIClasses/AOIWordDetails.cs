using EM_Analyzer.ExcelLogger;
using EM_Analyzer.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EM_Analyzer.ModelClasses.AOIClasses
{
    public class AOIWordDetails : IAOI
    {
        public static Dictionary<string, AOIWordDetails> nameToAOIWordsDetailsDictionary = new Dictionary<string, AOIWordDetails>();
        public static Dictionary<int, string> groupWordToSpecialName = new Dictionary<int, string>();

        private static bool includeFrequency = false;
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
        public int Length { get; set; }
        public int Frequency { get; set; }

        public AOIWordDetails(IEnumerable<string> details, uint lineNumber)
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

            // if there are 10 columns with frequencies or 9 columns without frequencies it means there is target name
            if ((count == 10 && includeFrequency) || (count == 9 && !includeFrequency))
            {
                SpecialNmae = enumerator.Current;
                groupWordToSpecialName[Group] = SpecialNmae;
                enumerator.MoveNext();
            }

            try
            {
                Length = int.Parse(enumerator.Current);
            }
            catch (Exception)
            {

                throw;
            }
            enumerator.MoveNext();
            try
            {
                if (includeFrequency)
                    Frequency = int.Parse(enumerator.Current);
            }
            catch (Exception)
            {

                throw;
            }

            if (IsProper)
            {
                DictionaryKey = Name + DictionaryKey;
                nameToAOIWordsDetailsDictionary[DictionaryKey] = this;
                AOIsService.nameToAOIWordsDictionary[DictionaryKey] = this;
            }
        }

        public static void LoadAllAOIWordFromFile(string fileName)
        {
            List<IEnumerable<string>> table = ExcelsService.ReadExcelFile<string>(fileName);
            uint lineNumber = 0;
            foreach (IEnumerable<string> details in table)
            {
                // in headres line of excel check if there is Frequency column
                if (lineNumber == 0)
                {
                    if (details.Contains("Frequency"))
                        includeFrequency = true;
                    lineNumber++;
                    continue;
                }

                List<string> detailsStr = details.ToList();
                new AOIWordDetails(detailsStr, lineNumber);
                lineNumber++;
                var T = AOIsService.nameToAOIWordsDictionary;

            }
        }
        private static Log CreateLogForFieldValidation(string fieldName, string valueFound, uint lineNumber)
        {
            return new Log() { FileName = FixationsService.wordsExcelFileName, LineNumber = lineNumber, Description = "The Value Of Field " + fieldName + " Is Not Valid!!!" + Environment.NewLine + "The Value Found Is: " + valueFound };
        }

        public double DistanceToAOI(Fixation fixation)
        {
            throw new NotImplementedException();
        }
    }
}