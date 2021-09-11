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
using EM_Analyzer.ExcelLogger;
using OfficeOpenXml;

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

            uint lineNumber = 1;
            foreach (Fixation fixation in table)
            {
                if (fixation.AOI_Name > 0)
                {
                    if (fixation.AOI_Phrase_Details.DistanceToAOI(fixation) != 0)
                    {
                        ExcelLoggerService.AddLog(new Log() { FileName = ConfigurationService.FirstExcelFileName, LineNumber = lineNumber+1, Description = "The Fixation Is Not Inside The AOI Name: " + fixation.AOI_Name });
                    }
                }
                lineNumber++;
            }

            ExcelsService.CreateExcelFromStringTable(ConfigurationService.FirstExcelFileName, table, EditExcelForFirstFile);
        }
        public static int EditExcelForFirstFile(ExcelWorksheet ws)
        {

            for (int i = 2; i <= ws.Dimension.Rows; i++)
            {
                // change figure AOI group label from 0 to "figure"
                if (ws.Cells[i, Constans.firstFileWordIndexCol].Value.ToString() == "0")
                    ws.Cells[i, Constans.firstFileWordIndexCol].Value = "figure";

            }
            return 0;
        }
    }
}
