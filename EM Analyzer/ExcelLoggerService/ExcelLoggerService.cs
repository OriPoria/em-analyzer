using EM_Analyzer.ModelClasses;
using EM_Analyzer.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EM_Analyzer.ExcelLogger
{
    public static class ExcelLoggerService
    {
        
        private static List<Log> logs = new List<Log>();
        static ExcelLoggerService()
        {
            AppDomain.CurrentDomain.ProcessExit += new EventHandler(OnProgramExit);
        }

        private static void OnProgramExit(object sender, EventArgs e)
        {
            ExcelsService.CreateExcelFromStringTable(ConfigurationService.LogsExcelFileName, logs);
        }

        public static void AddLog(Log log)
        {
            lock (logs)
            {
                logs.Add(log);
            }
        }
    }
}
