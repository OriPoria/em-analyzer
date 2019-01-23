using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace EM_Analyzer.ModelClasses
{
    class ConfigurationService
    {
        private const string CONFIG_FILE = "./Configuration.xml";
        public const string Minimum_Event_Duration_In_ms = "Minimum_Event_Duration_In_ms";
        public const string Number_Of_Fixations_In_AOI_For_Exception = "Number_Of_Fixations_In_AOI_For_Exception";
        public const string Number_Of_Fixations_Out_Of_AOI_For_Exception = "Number_Of_Fixations_Out_Of_AOI_For_Exception";
        public const string First_Excel_File_Name = "First_Excel_File_Name";
        public const string Second_Excel_File_Name = "Second_Excel_File_Name";
        public const string Third_Excel_File_Name = "Third_Excel_File_Name";
        public const string Dealing_With_Exceptions_Limit_In_Pixels = "Dealing_With_Exceptions_Limit_In_Pixels";
        public const string Dealing_With_Exceptions_Inside_The_Limit = "Dealing_With_Exceptions_Inside_The_Limit";
        public const string Dealing_With_Exceptions_Outside_The_Limit = "Dealing_With_Exceptions_Outside_The_Limit";

        static XmlDocument doc = new XmlDocument();

        static ConfigurationService()
        {
            doc.Load(CONFIG_FILE);
        }

        public static string getValue(string TagName)
        {
            return doc.GetElementsByTagName(TagName)[0].InnerText.Trim();
        }

        public static string FirstExcelFileName { get { return getValue(First_Excel_File_Name); } }
        public static string SecondExcelFileName { get { return getValue(Second_Excel_File_Name); } }
        public static string ThirdExcelFileName { get { return getValue(Third_Excel_File_Name); } }
        public static string MinimumEventDurationInms { get { return getValue(Minimum_Event_Duration_In_ms); } }
        public static string NumberOfFixationsInAOIForException { get { return getValue(Number_Of_Fixations_In_AOI_For_Exception); } }
        public static string NumberOfFixationsOutOfAOIForException { get { return getValue(Number_Of_Fixations_Out_Of_AOI_For_Exception); } }
        public static string DealingWithExceptionsLimitInPixels { get { return getValue(Dealing_With_Exceptions_Limit_In_Pixels); } }
        public static string DealingWithExceptionsInsideTheLimit { get { return getValue(Dealing_With_Exceptions_Inside_The_Limit); } }
        public static string DealingWithExceptionsOutsideTheLimit { get { return getValue(Dealing_With_Exceptions_Outside_The_Limit); } }
    }
}
