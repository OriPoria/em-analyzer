using EM_Analyzer.Enums;
using System.Collections.Generic;

namespace EM_Analyzer.ExcelsFilesMakers
{
    public class Constans
    {
        public static int secondFileStartCondsInx = 6;
        public static int secondFilePossibleConds = 35;
        public static int secondFileAoiTargetCol = 5;
        public static int secondFileAoiGroupCol = 4;
        public static int firstFileWordIndexCol = 4;
        public static int firstFileAOINameCol = 5;
        public static int firstFileAOIGroupBeforeChangeCol = 6;
        public static int firstFileAOIGroupAfterChangeCol = 7;


        public static List<string> parseSpecialName(string s)
        {
            string[] subs = s.Split('/');
            List<string> names = new List<string>();
            foreach (string str in subs)
                names.Add(str);
            return names;
        }
        public static string GetEndOfFileNameByType(AOITypes type)
        {
            if (type == AOITypes.Phrases)
                return "c";
            return "w";
        }

    }
}
