using System.Xml;

namespace EM_Analyzer.ModelClasses
{
    class ConfigurationService
    {
        public const string CONFIG_FILE = "./Configuration.xml";
        public const string Minimum_Event_Duration_In_ms = "Minimum_Event_Duration_In_ms";
        public const string Maximum_Event_Duration_In_ms = "Maximum_Event_Duration_In_ms";
        
        public const string Number_Of_Fixations_Out_AOI_For_Exception = "Number_Of_Fixations_Out_Of_AOI_For_Exception";
        public const string Number_Of_Fixations_In_Of_AOI_For_Exception = "Number_Of_Fixations_In_AOI_For_Exception";
        
        public const string Minimum_Number_Of_Fixations_For_Regression = "Minimum_Number_Of_Fixations_For_Regression";
        public const string Minimum_Duration_Of_Fixation_For_Regression = "Minimum_Duration_Of_Fixation_For_Regression";

        public const string Minimum_Number_Of_Fixations_For_First_Pass = "Minimum_Number_Of_Fixations_For_First_Pass";
        public const string Minimum_Duration_Of_Fixation_For_First_Pass = "Minimum_Duration_Of_Fixation_For_First_Pass";

        public const string Minimum_Number_Of_Fixations_For_Skip = "Minimum_Number_Of_Fixations_For_Skip";
        public const string Minimum_Duration_Of_Fixation_For_Skip = "Minimum_Duration_Of_Fixation_For_Skip";

        public const string Minimum_Number_Of_Fixations_For_Page_Visit = "Minimum_Number_Of_Fixations_For_Page_Visit";
        public const string Minimum_Duration_Of_Fixation_For_Page_Visit = "Minimum_Duration_Of_Fixation_For_Page_Visit";

        public const string First_Excel_File_Name = "First_Excel_File_Name";//Considered_Second_Excel_File_Name
        public const string Second_Excel_File_Name = "Second_Excel_File_Name";
        public const string Considered_Second_Excel_File_Name = "Considered_Second_Excel_File_Name";
        public const string Third_Excel_File_Name = "Third_Excel_File_Name";
        public const string Fourth_Excel_File_Name = "Fourth_Excel_File_Name";
        public const string Dealing_With_Exceptions_Limit_In_Pixels = "Dealing_With_Exceptions_Limit_In_Pixels";
        public const string Dealing_With_Exceptions_Inside_The_Limit = "Dealing_With_Exceptions_Inside_The_Limit";
        public const string Dealing_With_Exceptions_Outside_The_Limit = "Dealing_With_Exceptions_Outside_The_Limit";
        public const string Remove_Fixations_Appeared_Before_First_AOI = "Remove_Fixations_Appeared_Before_First_AOI";
        public const string Number_Of_Fixations_To_Remove_Before_First_AOI = "Number_Of_Fixations_To_Remove_Before_First_AOI";
        public const string Standard_Deviation = "Standard_Deviation";
        public const string Excel_Files_Extension = "Excel_Files_Extension";
        public const string Logs_Excel_File_Name = "Logs_Excel_File_Name";

        
        public const string Second_File_Filtering_denominator = "Second_File_Filtering_denominator";

        public const string Analyze_Extent = "Analyze_Extent";

        public const string Preview_Limit = "Preview_Limit";
        public const string Fixed_Time_In_sec = "Fixed_Time_In_sec";
        public const string Percentage_of_Participant_Time = "Percentage_of_Participant_Time";


        static XmlDocument doc = new XmlDocument();

        static ConfigurationService()
        {
            doc.Load(CONFIG_FILE);
        }

        public static string GetValue(string TagName)
        {
            return doc.GetElementsByTagName(TagName)[0].InnerText.Trim();
        }
        public static string FirstExcelFileName { get { return GetValue(First_Excel_File_Name); } }
        public static string SecondExcelFileName { get { return GetValue(Second_Excel_File_Name); } }
        public static string ConsideredSecondExcelFileName { get { return GetValue(Considered_Second_Excel_File_Name); } }
        public static string ThirdExcelFileName { get { return GetValue(Third_Excel_File_Name); } }
        public static string FourthExcelFileName { get { return GetValue(Fourth_Excel_File_Name); } }
        public static string MinimumEventDurationInms { get { return GetValue(Minimum_Event_Duration_In_ms); } }
        public static string MaximumEventDurationInms { get { return GetValue(Maximum_Event_Duration_In_ms); } }
        public static string MinimumNumberOfFixationsForRegression { get { return GetValue(Minimum_Number_Of_Fixations_For_Regression); } }
        public static string MinimumDurationOfFixationForRegression { get { return GetValue(Minimum_Duration_Of_Fixation_For_Regression); } }
        public static string MinimumNumberOfFixationsForFirstPass { get { return GetValue(Minimum_Number_Of_Fixations_For_First_Pass); } }
        public static string MinimumDurationOfFixationForFirstPass { get { return GetValue(Minimum_Duration_Of_Fixation_For_First_Pass); } }
        public static string MinimumNumberOfFixationsForSkip { get { return GetValue(Minimum_Number_Of_Fixations_For_Skip); } }
        public static string MinimumDurationOfFixationForSkip { get { return GetValue(Minimum_Duration_Of_Fixation_For_Skip); } }
        public static string MinimumNumberOfFixationsForPageVisit { get { return GetValue(Minimum_Number_Of_Fixations_For_Page_Visit); } }
        public static string MinimumDurationOfFixationForPageVisit { get { return GetValue(Minimum_Duration_Of_Fixation_For_Page_Visit); } }


        public static string NumberOfFixationsOutAOIForException { get { return GetValue(Number_Of_Fixations_Out_AOI_For_Exception); } }
        public static string NumberOfFixationsInOfAOIForException { get { return GetValue(Number_Of_Fixations_In_Of_AOI_For_Exception); } }
        public static string DealingWithExceptionsLimitInPixels { get { return GetValue(Dealing_With_Exceptions_Limit_In_Pixels); } }
        public static string DealingWithExceptionsInsideTheLimit { get { return GetValue(Dealing_With_Exceptions_Inside_The_Limit); } }
        public static string DealingWithExceptionsOutsideTheLimit { get { return GetValue(Dealing_With_Exceptions_Outside_The_Limit); } }
        public static string RemoveFixationsAppearedBeforeFirstAOI { get { return GetValue(Remove_Fixations_Appeared_Before_First_AOI); } }
        public static string NumberOfFixationsToRemoveBeforeFirstAOI { get { return GetValue(Number_Of_Fixations_To_Remove_Before_First_AOI); } }
        public static string StandardDeviation { get { return GetValue(Standard_Deviation); } }
        public static string ExcelFilesExtension { get { return GetValue(Excel_Files_Extension); } }
        public static string LogsExcelFileName { get { return GetValue(Logs_Excel_File_Name); } }
        public static string SecondFileFilteringDenominator { get { return GetValue(Second_File_Filtering_denominator); } }
        public static string AnalyzeExtent { get { return GetValue(Analyze_Extent); } }
        public static string PreviewLimit { get { return GetValue(Preview_Limit); } }
        public static string FixedTimeInSec { get { return GetValue(Fixed_Time_In_sec); } }
        public static string PercentageOfParticipantTime { get { return GetValue(Percentage_of_Participant_Time); } }

    }
}
