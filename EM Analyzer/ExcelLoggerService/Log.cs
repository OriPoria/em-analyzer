using System.ComponentModel;

namespace EM_Analyzer.ExcelLogger
{
    public class Log
    {
        //[XLColumn(Header ="File Name")]
        public string FileName { get; set; }
        [Description("Line")]
        public uint LineNumber { get; set; }
        [Description("Description")]
        public string Description { get; set; }
    }
}
