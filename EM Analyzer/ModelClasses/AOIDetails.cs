using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EM_Analyzer.Enums;
using EM_Analyzer.ExcelLogger;
using EM_Analyzer.Services;

namespace EM_Analyzer.ModelClasses
{
    class AOIDetails
    {
        public static Dictionary<string, AOIDetails> nameToAOIDetailsDictionary = new Dictionary<string, AOIDetails>();
        public static bool isAOIIncludeStimulus=false;
        public string Stimulus { get; set; }
        public int Name { get; set; }
        public int Group { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double H { get; set; }
        public double L { get; set; }
        public bool isProper { get; set; }

        public AOIDetails(IEnumerable<string> details, uint lineNumber)
        {
            this.isProper = true;
            IEnumerator<string> enumerator = details.GetEnumerator();
            string dictionaryKey="";

            enumerator.MoveNext();
            int count = details.Count();
            if (count >= 7)
            {
                this.Stimulus = enumerator.Current;
                enumerator.MoveNext();
                dictionaryKey = this.Stimulus;
            }
            if(count>=6)
            {
                try
                {
                    this.Name = int.Parse(enumerator.Current);
                }
                catch
                {
                    ExcelLoggerService.AddLog(CreateLogForFieldValidation("Name", enumerator.Current, lineNumber));
                    this.isProper = false;
                }
                enumerator.MoveNext();

                try
                {
                    this.Group = int.Parse(enumerator.Current);
                }
                catch
                {
                    ExcelLoggerService.AddLog(CreateLogForFieldValidation("Group", enumerator.Current, lineNumber));
                    this.isProper = false;
                }
                enumerator.MoveNext();

                try
                {
                    this.X = double.Parse(enumerator.Current);
                }
                catch
                {
                    ExcelLoggerService.AddLog(CreateLogForFieldValidation("X", enumerator.Current, lineNumber));
                    this.isProper = false;
                }
                enumerator.MoveNext();

                try
                {
                    this.Y = double.Parse(enumerator.Current);
                }
                catch
                {
                    ExcelLoggerService.AddLog(CreateLogForFieldValidation("Y", enumerator.Current, lineNumber));
                    this.isProper = false;
                }
                enumerator.MoveNext();

                try
                {
                    this.H = double.Parse(enumerator.Current);
                }
                catch
                {
                    ExcelLoggerService.AddLog(CreateLogForFieldValidation("H", enumerator.Current, lineNumber));
                    this.isProper = false;
                }
                enumerator.MoveNext();

                try
                {
                    this.L = double.Parse(enumerator.Current);
                }
                catch
                {
                    ExcelLoggerService.AddLog(CreateLogForFieldValidation("L", enumerator.Current, lineNumber));
                    this.isProper = false;
                }
                enumerator.MoveNext();

                if (this.isProper)
                {
                    dictionaryKey = this.Name + dictionaryKey;
                    nameToAOIDetailsDictionary[dictionaryKey] = this;
                }
            }
            /*
            {
                this.Name = int.Parse(details[(int)ExcelAOITableEnum.Name]);
                this.Group = int.Parse(details[(int)ExcelAOITableEnum.Group]);
                this.X = double.Parse(details[(int)ExcelAOITableEnum.X]);
                this.Y = double.Parse(details[(int)ExcelAOITableEnum.Y]);
                this.H = double.Parse(details[(int)ExcelAOITableEnum.H]);
                this.L = double.Parse(details[(int)ExcelAOITableEnum.L]);
                nameToAOIDetailsDictionary[this.Name+ this.Stimulus] = this;
            }
            */
        }

        public static void LoadAllAOIFromFile(string fileName)
        {
            List<IEnumerable<string>> table = ExcelsService.ReadExcelFile<string>(fileName);
            isAOIIncludeStimulus = false;
            if (table.FirstOrDefault()?.Count() >= 7)
                isAOIIncludeStimulus = true;
            //table.AsParallel().(details => new AOIDetails(details));
            uint lineNumber = 1;
            foreach(IEnumerable<string> details in table)
            {
                new AOIDetails(details, lineNumber);
                lineNumber++;
            }
        }

        private static Log CreateLogForFieldValidation(string fieldName, string valueFound, uint lineNumber)
        {
            return new Log() { FileName = FixationsService.excelFileName, LineNumber = lineNumber, Description = "The Value Of Field " + fieldName + " Is Not Valid!!!" + Environment.NewLine + "The Value Found Is: " + valueFound };
        }
    }
}
