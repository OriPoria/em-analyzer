using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EM_Analyzer.Interfaces
{
    interface ITrialClassForConsideringCoverage
    {


        [Description("Participant")]
        string Participant { get; }
        [Description("Trial")]
        string Trial { get; }
        [Description("Stimulus")]
        string Stimulus { get; }
        [Description("Mean Fixation Duration")]
        double Mean_Fixation_Duration { get;  }
        [Description("SD Fixation Duration")]
        double SD_Fixation_Duration { get; }
        [Description("Total Fixation Number")]
        int Total_Fixation_Number { get;  }
        [Description("Mean Progressive Fixation Duration")]
        double Mean_Progressive_Fixation_Duration { get; }
        [Description("SD Progressive Fixation Duration")]
        double SD_Progressive_Fixation_Duration { get; }
        [Description("Progressive Fixation Number")]
        int Progressive_Fixation_Number { get;  }
        [Description("Mean Progressive Saccade Length")]
        double Mean_Progressive_Saccade_Length { get;  }
        [Description("SD Progressive Saccade Length")]
        double SD_Progressive_Saccade_Length { get;  }
        [Description("Mean Progressive Saccade Length X")]
        double Mean_Progressive_Saccade_Length_X { get; }
        [Description("SD Progressive Saccade Length X")]
        double SD_Progressive_Saccade_Length_X { get; }
        [Description("Mean Regressive Fixation Duration")]
        double Mean_Regressive_Fixation_Duration { get; }
        [Description("SD Regressive Fixation Duration")]
        double SD_Regressive_Fixation_Duration { get; }
        [Description("Regressive Fixation Number")]
        int Regressive_Fixation_Number { get; }
        [Description("Mean Regressive Saccade Length")]
        double Mean_Regressive_Saccade_Length { get; }
        [Description("SD Regressive Saccade Length")]
        double SD_Regressive_Saccade_Length { get; }
        [Description("Mean Regressive Saccade Length X")]
        double Mean_Regressive_Saccade_Length_X { get; }
        [Description("SD Regressive Saccade Length X")]
        double SD_Regressive_Saccade_Length_X { get; }
        [Description("Mean Pupil Diameter [mm]")]
        double Mean_Pupil_Diameter { get; }
    }
}
