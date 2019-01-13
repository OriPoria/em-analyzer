using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ForBarIlanResearch.Enums;
using ForBarIlanResearch.Services;

namespace ForBarIlanResearch.ModelClasses
{
    class AOIDetails
    {
        public static Dictionary<string, AOIDetails> nameToAOIDetailsDictionary = new Dictionary<string, AOIDetails>();
        public string Stimulus { get; set; }
        public int Name { get; set; }
        public int Group { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double H { get; set; }
        public double L { get; set; }

        public AOIDetails(string[] details)
        {
            if(details.Length>=7)
            {
                this.Stimulus = details[(int)ExcelAOITableEnum.Stimulus];
                this.Name = int.Parse(details[(int)ExcelAOITableEnum.Name]);
                this.Group = int.Parse(details[(int)ExcelAOITableEnum.Group]);
                this.X = double.Parse(details[(int)ExcelAOITableEnum.X]);
                this.Y = double.Parse(details[(int)ExcelAOITableEnum.Y]);
                this.H = double.Parse(details[(int)ExcelAOITableEnum.H]);
                this.L = double.Parse(details[(int)ExcelAOITableEnum.L]);
                nameToAOIDetailsDictionary[this.Name+ this.Stimulus] = this;
            }
        }

        public static void LoadAllAOIFromFile(string fileName)
        {
            // List<double[]> table= ExcelsService.ReadExcelFile<double>(fileName);
            List<string[]> table = ExcelsService.ReadExcelFile<string>(fileName);

            table.ForEach(details => new AOIDetails(details));
        }
    }
}
