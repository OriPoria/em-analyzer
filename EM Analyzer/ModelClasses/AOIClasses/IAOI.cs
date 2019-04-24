using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EM_Analyzer.ModelClasses;

namespace EM_Analyzer.ModelClasses.AOIClasses
{
    public interface IAOI
    {
        double DistanceToAOI(Fixation fixation);
        double AOI_Coverage_In_Percents { get; set; }
        double AOI_Size_X { get; set; }
        //bool IsProper { get; set; }
        int Group { get; set; }
    }
}
