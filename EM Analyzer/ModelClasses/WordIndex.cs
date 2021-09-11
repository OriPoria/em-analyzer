using EM_Analyzer.Enums;
using EM_Analyzer.ExcelLogger;
using EM_Analyzer.ModelClasses;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EM_Analyzer
{
    public class WordIndex
    {
        public long Index { get; set; }
        public int Group{ get; set; }
        public long AOI_Word_Size { get; set; }
        public double AOI_Coverage_In_Percents { get; set; }
        
        static readonly double minimumDuration = double.Parse(ConfigurationService.MinimumEventDurationInms);
        static readonly double maximumDuration = double.Parse(ConfigurationService.MaximumEventDurationInms);
        public static WordIndex CreateWordIndexFromStringArray(string[] arr, uint lineNumber)
        {
            Fixation newFixation = new Fixation
            {
                Trial = arr[TextFileColumnIndexes.Trial].Trim(),
                Stimulus = arr[TextFileColumnIndexes.Stimulus].Trim(),
                Participant = arr[TextFileColumnIndexes.Participant].Trim()
            };
            WordIndex wordIndex = new WordIndex();
            bool isValid = true;
            try
            {
                newFixation.Event_Duration = double.Parse(arr[TextFileColumnIndexes.Event_Duration]);
                if (newFixation.Event_Duration < minimumDuration || newFixation.Event_Duration > maximumDuration)
                    return null;
            }

            catch
            {
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("Event Duration", arr[TextFileColumnIndexes.Event_Duration], lineNumber));
                isValid = false;
            }

            try
            {
                wordIndex.Group = int.Parse(arr[TextFileColumnIndexes.AOI_Group]);
            }
            catch
            {
                wordIndex.Group = -1;
            }
            try
            {
                wordIndex.Index = long.Parse(arr[TextFileColumnIndexes.Index]);
            }
            catch
            {
                isValid = false;
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("Index", arr[TextFileColumnIndexes.Index], lineNumber));
            }
            try
            {
                wordIndex.AOI_Word_Size = long.Parse(arr[TextFileColumnIndexes.AOI_Size]);
            }
            catch (Exception)
            {
                isValid = false;
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("AOI_Size", arr[TextFileColumnIndexes.AOI_Size], lineNumber));
            }
            try
            {
                wordIndex.AOI_Coverage_In_Percents = double.Parse(arr[TextFileColumnIndexes.AOI_Coverage]);
            }
            catch
            {
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("AOI_Coverage", arr[TextFileColumnIndexes.AOI_Coverage], lineNumber));
                isValid = false;
            }
            if (!isValid)
                return null;

            string dictionatyKey = newFixation.GetDictionaryKey();
            if (!FixationsService.wordIndexSetToFixationListDictionary.ContainsKey(dictionatyKey))
                FixationsService.wordIndexSetToFixationListDictionary[dictionatyKey] = new List<WordIndex>();
            FixationsService.wordIndexSetToFixationListDictionary[dictionatyKey].Add(wordIndex);

            return wordIndex;

        }
        private static Log CreateLogForFieldValidation(string fieldName, string valueFound, uint lineNumber)
        {
            return new Log() { FileName = FixationsService.wordsTextFileName, LineNumber = lineNumber, Description = "The Value Of Field " + fieldName + " Is Not Valid!!! " + Environment.NewLine + "The Value Found Is: " + valueFound };
        }
    }
}
