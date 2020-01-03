using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EM_Analyzer.Interfaces
{
    public interface IAOIClassForConsideringCoverage
    {
        [Description("Participant")]
        string Participant { get; }
        [Description("Trial")]
        string Trial { get; }
        [Description("Stimulus")]
        string Stimulus { get; }
        [Description("Text Name")]
        string Text_Name { get; }
        [Description("AOI Group")]
        int AOI_Group { get; }
        [Description("Total Fixation Duration")]
        double Total_Fixation_Duration { get; }
        [Description("Total Fixation Number")]
        double Total_Fixation_Number { get; }
        [Description("First Fixation Duration")]
        double First_Fixation_Duration { get; }
        [Description("First-Pass Duration")]
        double First_Pass_Duration { get; }
        [Description("First-Pass Number")]
        double First_Pass_Number { get; }
        [Description("First-Pass Progressive Duration")]
        double First_Pass_Progressive_Duration { get; }
        [Description("First-Pass Progressive Number")]
        double First_Pass_Progressive_Number { get; }
        [Description("First-Pass Progressive Duration Overall")]
        double First_Pass_Progressive_Duration_Overall { get; }
        [Description("First-Pass Progressive Number Overall")]
        double First_Pass_Progressive_Number_Overall { get; }
        [Description("Total First-Pass Progressive Duration")]
        double Total_First_Pass_Progressive_Duration { get; }
        [Description("Total First-Pass Progressive Number")]
        double Total_First_Pass_Progressive_Number { get; }
        [Description("Total First-Pass Progressive Duration Overall")]
        double Total_First_Pass_Progressive_Duration_Overall { get; }
        [Description("Total First-Pass Progressive Number Overall")]
        double Total_First_Pass_Progressive_Number_Overall { get; }
        [Description("Total First-Pass Regressive Duration")]
        double Total_First_Pass_Regressive_Duration { get; }
        [Description("Total First-Pass Regressive Number")]
        double Total_First_Pass_Regressive_Number { get; }
        [Description("Regression Number")]
        double Regression_Number { get; }
        [Description("Regression Duration")]
        double Regression_Duration { get; }
        [Description("First Regression Duration")]
        double First_Regression_Duration { get; }
        [Description("Skip")]
        bool Skip { get; set; }
        [Description("Pupil Diameter [mm]")]
        double Pupil_Diameter { get; }
        [Description("AOI Size X [mm]")]
        double Mean_AOI_Size { get; }
        [Description("AOI Coverage [%]")]
        double Mean_AOI_Coverage { get; }
    }
}
