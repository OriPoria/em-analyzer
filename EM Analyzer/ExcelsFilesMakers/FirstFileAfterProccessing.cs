using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using EM_Analyzer.ModelClasses;
using EM_Analyzer.Enums;
using EM_Analyzer.Services;
using System.Threading;

namespace EM_Analyzer.ExcelsFilesMakers
{
    class FirstFileAfterProccessing
    {
        public static void MakeExcelFile()
        {
            List<Fixation> table = new List<Fixation>();
            List<Fixation>[] values = FixationsService.fixationSetToFixationListDictionary.Values.ToArray();
            foreach (List<Fixation> fixations in values)
                table.AddRange(fixations);
            ExcelsService.CreateExcelFromStringTable(ConfigurationService.FirstExcelFileName, table);
        }
    }
}
