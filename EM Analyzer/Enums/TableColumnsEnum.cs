using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EM_Analyzer.Enums
{
    //enum TableColumnsEnum : int
    //{
    //    Trial,
    //    Stimulus,
    //    Export_Start_Trial_Time,
    //    Export_End_Trial_Time,
    //    Participant,
    //    Color,
    //    Category_Group,
    //    Category,
    //    Eye_L_R,
    //    AOI_Name,
    //    AOI_Group,
    //    AOI_Scope,
    //    AOI_Order,
    //    AOI_Size,
    //    AOI_Coverage,
    //    Time_to_First_Appearance,
    //    Appearance_Count,
    //    Visible_Time_in_ms,
    //    Visible_Time_in_percents,
    //    Index,
    //    Event_Start_Trial_Time,
    //    Event_End_Trial_Time,
    //    Event_Duration,
    //    Fixation_Position_X,
    //    Fixation_Position_Y,
    //    Fixation_Average_Pupil_Size_X,
    //    Fixation_Average_Pupil_Size_Y,
    //    Fixation_Average_Pupil_Diameter,
    //    Fixation_Dispersion_X,
    //    Fixation_Dispersion_Y,
    //    Mouse_Position_X,
    //    Mouse_Position_Y, 
    //    Is_Exception
    //}
    
    enum ExcelAOITableEnum
    {
        Stimulus,
        Name,
        Group,
        X,
        Y,
        H,
        L
    }

    enum DealingWithExceptionsEnum
    {
        Do_Nothing=1,
        Skip_In_First_Pass,
        Change_AOI_Group
    }

    enum DealingWithExceptionsOutBoundsEnum
    {
        Do_Nothing = 1,
        Skip_In_First_Pass=2
    }
    public enum AOITypes
    {
        Phrases = 1,
        Words = 2
    }

    // For the text file with the bad headers (not relevant for now, maybe for the future).
    static class TextFileColumnIndexes
    {
        public static int Trial,
        Stimulus,
        Participant,
        AOI_Name,
        AOI_Group,
        AOI_Coverage,
        Index,
        Event_Duration,
        Fixation_Position_X,
        Fixation_Position_Y,
        Fixation_Average_Pupil_Diameter,
        AOI_Size,
        Word_Index;
    }
}
