using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static EM_Analyzer.ExcelsFilesMakers.SecondFileConsideringCoverage;
using EM_Analyzer.Interfaces;

namespace EM_Analyzer.ExcelsFilesMakers
{
    public class AIOClassFromExcel : IAOIClassForConsideringCoverage
    {
        [Description("Participant")]
        public string Participant { get; set; }
        [Description("Trial")]
        public string Trial { get; set; }
        [Description("Stimulus")]
        public string Stimulus { get; set; }
        [Description("Text Name")]
        public string Text_Name { get; set; }
        [Description("AOI Group")]
        public int AOI_Group { get; set; }
        [Description("AOI Target")]
        public string AOI_Target { get; set; }
        public List<string> List_Of_Strings { get; set; }

        [Description("Total Fixation Duration")]
        public double Total_Fixation_Duration { get; set; }

        [Description("Total Fixation Number")]
        public double Total_Fixation_Number { get; set; }

        [Description("First Fixation Duration")]
        public double First_Fixation_Duration { get; set; }

        [Description("First-Pass Duration")]
        public double First_Pass_Duration { get; set; }

        [Description("First-Pass Number")]
        public double First_Pass_Number { get; set; }

        [Description("First-Pass Progressive Duration")]
        public double First_Pass_Progressive_Duration { get; set; }

        [Description("First-Pass Progressive Number")]
        public double First_Pass_Progressive_Number { get; set; }

        [Description("First-Pass Progressive Duration Overall")]
        public double First_Pass_Progressive_Duration_Overall { get; set; }

        [Description("First-Pass Progressive Number Overall")]
        public double First_Pass_Progressive_Number_Overall { get; set; }

        [Description("Total First-Pass Progressive Duration")]
        public double Total_First_Pass_Progressive_Duration { get; set; }

        [Description("Total First-Pass Progressive Number")]
        public double Total_First_Pass_Progressive_Number { get; set; }

        [Description("Total First-Pass Progressive Duration Overall")]
        public double Total_First_Pass_Progressive_Duration_Overall { get; set; }

        [Description("Total First-Pass Progressive Number Overall")]
        public double Total_First_Pass_Progressive_Number_Overall { get; set; }

        [Description("Total First-Pass Regressive Duration")]
        public double Total_First_Pass_Regressive_Duration { get; set; }

        [Description("Total First-Pass Regressive Number")]
        public double Total_First_Pass_Regressive_Number { get; set; }

        [Description("Regression Number")]
        public double Regression_Number { get; set; }

        [Description("Regression Duration")]
        public double Regression_Duration { get; set; }

        [Description("First Regression Duration")]
        public double First_Regression_Duration { get; set; }

        [Description("Skip")]
        public bool Skip { get; set; }

        [Description("Pupil Diameter [mm]")]
        public double Pupil_Diameter { get; set; }

        [Description("AOI Size X [mm]")]
        public double Mean_AOI_Size { get; set; }

        [Description("AOI Coverage [%]")]
        public double Mean_AOI_Coverage { get; set; }
        public int Length { get; }
        public int Frequency { get; }
        public void CreateAIOClassAfterCoverage()
        {
            new AIOClassAfterCoverage(this);
        }

        public AIOClassFromExcel()
        {

        }

    }
}
