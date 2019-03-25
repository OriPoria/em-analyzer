using ClosedXML.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EM_Analyzer.ExcelLogger
{
    public class Log
    {
        [XLColumn(Header ="File Name")]
        public string FileName { get; set; }
        [XLColumn(Header = "Line")]
        public uint LineNumber { get; set; }
        [XLColumn(Header = "Description")]
        public string Description { get; set; }
    }
}
