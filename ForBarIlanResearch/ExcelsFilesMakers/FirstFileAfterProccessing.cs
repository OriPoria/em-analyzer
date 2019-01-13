using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using ForBarIlanResearch.ModelClasses;
using ForBarIlanResearch.Enums;
using ForBarIlanResearch.Services;
using System.Threading;

namespace ForBarIlanResearch.ExcelsFilesMakers
{
    class FirstFileAfterProccessing
    {
        //public static Thread makeExcelFile()
        public static void makeExcelFile()
        {
            List<Fixation> table = new List<Fixation>();
            List<Fixation>[] values = FixationsService.fixationSetToFixationListDictionary.Values.ToArray();
            //Thread makeFirstFile = new Thread(() =>
            //  {
                  foreach (List<Fixation> fixations in values)
                      table.AddRange(fixations);
                  ExcelsService.CreateExcelFromStringTable(ConfigurationService.getValue(ConfigurationService.First_Excel_File_Name), table);
            //  });
            //makeFirstFile.Start();
            //return makeFirstFile;
        }
    }
}
