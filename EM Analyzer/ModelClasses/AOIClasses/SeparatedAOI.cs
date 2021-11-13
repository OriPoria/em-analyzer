using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EM_Analyzer.ModelClasses.AOIClasses
{
    public class SeparatedAOI : IAOI
    {
        public double AOI_Coverage_In_Percents { get; set; }
        public double AOI_Size_X { get; set; }
        public int Group { get; set; }
        public string SpecialNmae { get; set; }

        private List<IAOI> AOIs = new List<IAOI>();

        public SeparatedAOI(IEnumerable<IAOI> AOIs)
        {
            this.AOIs.AddRange(AOIs);
            this.Group = AOIs.First().Group;
            this.AOI_Coverage_In_Percents = 0;
            this.AOI_Size_X = 0;
            foreach(var aoi in AOIs)
            {
                this.AOI_Coverage_In_Percents += aoi.AOI_Coverage_In_Percents;
                this.AOI_Size_X += aoi.AOI_Size_X;
            }
        }
        /*
        public double DistanceToAOI(Fixation fixation)
        {
            return this.AOIs.Min(aoi => aoi.DistanceToAOI(fixation));
        }
        */
    }
}
