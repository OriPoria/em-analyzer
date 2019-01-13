using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ForBarIlanResearch.ModelClasses
{
    /*
    class FixationSet
    {
        public string Participant { get; set; }
        public string Trial { get; set; }
        public string Stimulus { get; set; }
        public List<Fixation> Fixations { get; set; }
    }
    */
    /*
    class FixationSetComparer : IEqualityComparer<FixationSet>
    {
        public bool Equals(FixationSet x, FixationSet y)
        {
            if (x == null && y == null)
                return true;
            if (x == null || y == null)
                return false;
            return x.Fixations == y.Fixations;
        }

        public int GetHashCode(FixationSet obj)
        {
            return obj.Fixations.GetHashCode();
            //throw new NotImplementedException();
        }
    }
    */
}
