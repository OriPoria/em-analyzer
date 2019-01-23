using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EM_Analyzer.Enums;
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

        public AOIDetails(IEnumerable<string> details)
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
                this.Name = int.Parse(enumerator.Current);
                enumerator.MoveNext();
                this.Group = int.Parse(enumerator.Current);
                enumerator.MoveNext();
                this.X = double.Parse(enumerator.Current);
                enumerator.MoveNext();
                this.Y = double.Parse(enumerator.Current);
                enumerator.MoveNext();
                this.H = double.Parse(enumerator.Current);
                enumerator.MoveNext();
                this.L = double.Parse(enumerator.Current);
                enumerator.MoveNext();
                dictionaryKey = this.Name+dictionaryKey;
                nameToAOIDetailsDictionary[dictionaryKey] = this;
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
            // List<double[]> table= ExcelsService.ReadExcelFile<double>(fileName);
            List<IEnumerable<string>> table = ExcelsService.ReadExcelFile<string>(fileName);
            isAOIIncludeStimulus = false;
            if (table.FirstOrDefault()?.Count() >= 7)
                isAOIIncludeStimulus = true;
            table.ForEach(details => new AOIDetails(details));
        }
    }
}
