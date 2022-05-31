using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace EM_Analyzer
{
    public class StandardDevision
    {
        public static double ComputeStandardDevision(IEnumerable<double> values)
        {
            if (values.Count() == 0)
                return 0;
            double average = values.Average();
            int n = values.Count()-1;
            double standardDevision = Math.Sqrt(values.Sum(val => Math.Pow((val - average), 2)) / n);
            return standardDevision;
        }

        public static IEnumerable<double> ComputeStandardDevisionGrades(IEnumerable<double> values)
        {
            //List<double> list_of_values = values.ToList();
            //list_of_values.RemoveAll(item => double.IsInfinity(item));
            //values = list_of_values.AsEnumerable();
            double standardDevision = ComputeStandardDevision(values);
            double average = values.Average();
            IEnumerable<double> grades = values.Select(val => (val - average) / standardDevision);
            return grades;
        }
    }
}
