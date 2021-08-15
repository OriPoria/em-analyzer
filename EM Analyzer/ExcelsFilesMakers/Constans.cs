using EM_Analyzer.Enums;
using System.Collections.Generic;

namespace EM_Analyzer.ExcelsFilesMakers
{
    public class Constans
    {
        public static int startCondsInx = 7;
        public static int PossibleConds = 35;
        public static int aoiTargetCol = 6;

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
