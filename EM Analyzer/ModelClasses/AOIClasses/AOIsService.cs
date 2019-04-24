using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EM_Analyzer.Services;

namespace EM_Analyzer.ModelClasses.AOIClasses
{
    public static class AOIsService
    {
        public static Dictionary<string, IAOI> nameToAOIDictionary = new Dictionary<string, IAOI>();
        public static bool isAOIIncludeStimulus = false;
        
    }
}
