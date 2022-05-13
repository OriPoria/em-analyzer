using EM_Analyzer.ModelClasses;
using EM_Analyzer.Services;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;

namespace EM_Analyzer.ExcelsFilesMakers.ThirdFourFilter
{
    class ThirdFourthFilter
    {
        public static double standardDevisionAllowed;
        public delegate double NumericExpression(Fixation value);
        public static List<FilteredTrialTextPerParticipant> filteredTrialTextPerParticipants = new List<FilteredTrialTextPerParticipant>();

        /*
         All the dictionaries in this class are mapping between the trial and number of fixations that 
         are eliminated by the filter. 
         */
        public class FilteredTrialTextPerParticipant
        {
            public string Participant { get; set; }
            public string Stimulus { get; set; }
            public HashSet<string> All_Trials = new HashSet<string>(); // all valid trials 
            private List<Fixation> allFixations; // all fixations, include first in each trial
            public List<Fixation> m_Fixations_Text;
            private List<Fixation> m_Fixation_Filter_Duration;
            private Dictionary<string, int> m_Trial_Elimination_Fixations_Duration;
            public Dictionary<string, int> Trial_Elimination_Fixations_Duration 
            { 
                get
                {
                    if (this.m_Fixation_Filter_Duration == null)
                        this.m_Fixation_Filter_Duration = FilterFunction(m_Fixations_Text, out m_Trial_Elimination_Fixations_Duration, fix => fix.Event_Duration);
                    return m_Trial_Elimination_Fixations_Duration;
                }
            }
            public List<Fixation> All_Fixations_Duration_Filter
            {
                get
                {
                    if (this.m_Fixation_Filter_Duration == null)
                        this.m_Fixation_Filter_Duration = FilterFunction(m_Fixations_Text, out m_Trial_Elimination_Fixations_Duration, fix => fix.Event_Duration);
                    return m_Fixation_Filter_Duration;
                }
            }
            
            // progressive fixations
            public List<Fixation> m_Progressive_Fixations;
            private List<Fixation> m_Progressive_Fixations_Filter_Duration;
            private Dictionary<string, int> m_Trial_Elimination_Progressive_Duration;
            public Dictionary<string, int> Trial_Elimination_Progressive_Duration
            {
                get
                {
                    if (m_Progressive_Fixations_Filter_Duration == null)
                        this.m_Progressive_Fixations_Filter_Duration = FilterFunction(m_Progressive_Fixations, out m_Trial_Elimination_Progressive_Duration, fix => fix.Event_Duration);
                    return m_Trial_Elimination_Progressive_Duration;
                }
            }
            public List<Fixation> Progressive_Fixations_Duration_Filter
            { 
                get
                {
                    if (m_Progressive_Fixations_Filter_Duration == null)
                        this.m_Progressive_Fixations_Filter_Duration = FilterFunction(m_Progressive_Fixations, out m_Trial_Elimination_Progressive_Duration, fix => fix.Event_Duration);
                    return m_Progressive_Fixations_Filter_Duration;
                }
            }
            private List<Fixation> m_Progressive_Saccade_Length_Filter;
            private Dictionary<string, int> m_Trial_Elimination_Prog_Saccade_Length;
            public Dictionary<string, int> Trial_Elimination_Prog_Saccade_Length
            {
                get
                {
                    if (m_Progressive_Saccade_Length_Filter == null)
                        this.m_Progressive_Saccade_Length_Filter = FilterFunction(m_Progressive_Fixations, out m_Trial_Elimination_Prog_Saccade_Length, fix => fix.DistanceToPreviousFixation());
                    return m_Trial_Elimination_Prog_Saccade_Length;
                }
            }
            public List<Fixation> Progressive_Saccade_Length_Filter
            {
                get
                {
                    if (m_Progressive_Saccade_Length_Filter == null)
                        this.m_Progressive_Saccade_Length_Filter = FilterFunction(m_Progressive_Fixations, out m_Trial_Elimination_Prog_Saccade_Length, fix => fix.DistanceToPreviousFixation());
                    return m_Progressive_Saccade_Length_Filter;
                }
            }
            private List<Fixation> m_Progressive_Saccade_Length_X_Filter;
            private Dictionary<string, int> m_Trial_Elimination_Prog_Saccade_X_Length;
            public Dictionary<string, int> Trial_Elimination_Prog_Saccade_X_Length
            {
                get
                {
                    if (m_Progressive_Saccade_Length_X_Filter == null)
                        this.m_Progressive_Saccade_Length_X_Filter = FilterFunction(m_Progressive_Fixations, out m_Trial_Elimination_Prog_Saccade_X_Length, fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X));
                    return m_Trial_Elimination_Prog_Saccade_X_Length;
                }
            }
            public List<Fixation> Progressive_Saccade_Length_X_Filter
            {
                get
                {
                    if (m_Progressive_Saccade_Length_X_Filter == null)
                        this.m_Progressive_Saccade_Length_X_Filter = FilterFunction(m_Progressive_Fixations, out m_Trial_Elimination_Prog_Saccade_X_Length, fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X));
                    return m_Progressive_Saccade_Length_X_Filter;
                }
            }
            // regressive fixations
            public List<Fixation> m_Regressive_Fixations;
            private List<Fixation> m_Regressive_Fixations_Filter;
            private Dictionary<string, int> m_Trial_Elimination_Regressive_Duration;
            public Dictionary<string, int> Trial_Elimination_Regressive_Duration
            {
                get
                {
                    if (m_Regressive_Fixations_Filter == null)
                        this.m_Regressive_Fixations_Filter = FilterFunction(m_Regressive_Fixations, out m_Trial_Elimination_Regressive_Duration, fix => fix.Event_Duration);
                    return m_Trial_Elimination_Regressive_Duration;
                }
            }
            public List<Fixation> Regressive_Fixations_Duration_Filter
            {
                get
                {
                    if (m_Regressive_Fixations_Filter == null)
                        this.m_Regressive_Fixations_Filter = FilterFunction(m_Regressive_Fixations, out m_Trial_Elimination_Regressive_Duration, fix => fix.Event_Duration);
                    return m_Regressive_Fixations_Filter;
                }
            }
            private List<Fixation> m_Regressive_Saccade_Length_Filter;
            private Dictionary<string, int> m_Trial_Elimination_Reg_Saccade_Length;

            public Dictionary<string, int> Trial_Elimination_Reg_Saccade_Length
            {
                get
                {
                    if (m_Regressive_Saccade_Length_Filter == null)
                        this.m_Regressive_Saccade_Length_Filter = FilterFunction(m_Regressive_Fixations, out m_Trial_Elimination_Reg_Saccade_Length, fix => fix.DistanceToPreviousFixation());
                    return m_Trial_Elimination_Reg_Saccade_Length;
                }
            }
            public List<Fixation> Regressive_Saccade_Length_Filter
            {
                get
                {
                    if (m_Regressive_Saccade_Length_Filter == null)
                        this.m_Regressive_Saccade_Length_Filter = FilterFunction(m_Regressive_Fixations, out m_Trial_Elimination_Reg_Saccade_Length, fix => fix.DistanceToPreviousFixation());
                    return m_Regressive_Saccade_Length_Filter;

                }
            }
            private List<Fixation> m_Regressive_Saccade_Length_X_Filter;
            private Dictionary<string, int> m_Trial_Elimination_Reg_Saccade_X_Length;
            public Dictionary<string, int> Trial_Elimination_Reg_Saccade_X_Length
            {
                get
                {
                    if (m_Regressive_Saccade_Length_X_Filter == null)
                        this.m_Regressive_Saccade_Length_X_Filter  = FilterFunction(m_Regressive_Fixations, out m_Trial_Elimination_Reg_Saccade_X_Length, fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X));
                    return m_Trial_Elimination_Reg_Saccade_X_Length;
                }
            }
            public List<Fixation> Regressive_Saccade_Length_X_Filter
            {
                get
                {
                    if (m_Regressive_Saccade_Length_X_Filter == null)
                        this.m_Regressive_Saccade_Length_X_Filter = FilterFunction(m_Regressive_Fixations, out m_Trial_Elimination_Reg_Saccade_X_Length, fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X));
                    return m_Regressive_Saccade_Length_X_Filter;
                }
            }
            // Average_Pupil_Diameter
            private List<Fixation> m_Fixations_Average_Pupil_Diameter;
            private Dictionary<string, int> m_Trial_Elimination_Avg_Pupil_Diameter;
            public Dictionary<string, int> Trial_Elimination_Avg_Pupil_Diameter
            {
                get
                {
                    if (m_Fixations_Average_Pupil_Diameter == null)
                        m_Fixations_Average_Pupil_Diameter = FilterFunction(m_Fixations_Text, out m_Trial_Elimination_Avg_Pupil_Diameter, fix => fix.Fixation_Average_Pupil_Diameter);
                    return m_Trial_Elimination_Avg_Pupil_Diameter;
                }
            }
            public List<Fixation> Fixations_Average_Pupil_Diameter_Filter
            {
                get
                {
                    if (m_Fixations_Average_Pupil_Diameter == null)
                        m_Fixations_Average_Pupil_Diameter = FilterFunction(m_Fixations_Text, out m_Trial_Elimination_Avg_Pupil_Diameter, fix => fix.Fixation_Average_Pupil_Diameter);
                    return m_Fixations_Average_Pupil_Diameter;
                }
            }
            private List<int> m_Pages_Sequence;

            public int Page_Moves
            {
                get
                {
                    if (this.m_Pages_Sequence == null)
                    {
                        InitializePagesSequence();
                    }
                    return this.m_Pages_Sequence.Count - 1;
                }
            }
            public int Page_Regressions
            {
                get
                {
                    if (this.m_Pages_Sequence == null)
                    {
                        InitializePagesSequence();
                    }
                    int pageRegressions = 0;
                    int[] pages = this.m_Pages_Sequence.ToArray();
                    int prevPage = pages[0];
                    for (int i = 1; i < pages.Length; i++)
                    {
                        if (pages[i] < prevPage)
                            pageRegressions++;
                        prevPage = pages[i];
                    }
                    return pageRegressions;
                }
            }
            public int Page_Visits
            {
                get
                {
                    if (this.m_Pages_Sequence == null)
                    {
                        InitializePagesSequence();
                    }
                    return this.m_Pages_Sequence.Distinct().ToList().Count;
                }
            }


            public FilteredTrialTextPerParticipant (string Participant, string Stimulus, List<Fixation> fixations)
            {
                this.Stimulus = Stimulus;
                this.Participant = Participant;
                this.allFixations = fixations;

                // all fixations of text without the first one in every trial (and all the aoi's > 0)
                this.m_Fixations_Text = new List<Fixation>();
                List<IGrouping<string, Fixation>> groupingTrials = allFixations.GroupBy(x => x.Trial).ToList();
                foreach (IGrouping<string, Fixation> item in groupingTrials)
                    m_Fixations_Text.AddRange(item.Skip(1).ToList());
                this.m_Progressive_Fixations = null;
                this.m_Regressive_Fixations = null;
                InitializeProgressiveAndRegressiveFixationsList();

            }
            private void InitializeProgressiveAndRegressiveFixationsList()
            {
                this.m_Progressive_Fixations = new List<Fixation>();
                this.m_Regressive_Fixations = new List<Fixation>();
                Fixation[] fixations = this.allFixations.ToArray();
                int currentPage = fixations[0].Page;
                for (int i = 1; i < fixations.Length; ++i)
                {
                    if (currentPage != fixations[i].Page)
                    {
                        currentPage = fixations[i].Page;
                        continue;
                    }
                    All_Trials.Add(fixations[i].Trial); // create hash set of all Trials
                    fixations[i].Previous_Fixation = fixations[i - 1];
                    if (fixations[i - 1].IsBeforeThan(fixations[i]))
                        this.m_Progressive_Fixations.Add(fixations[i]);
                    else
                        this.m_Regressive_Fixations.Add(fixations[i]);
                }
            }
            private void InitializePagesSequence()
            {
                this.m_Pages_Sequence = new List<int>();
                Fixation[] fixations = m_Fixations_Text.ToArray();
                for (int i = 0; i < fixations.Length; i++)
                {
                    string currentTrial = fixations[i].Trial;
                    int currentPage = fixations[i].Page;
                    List<Fixation> pagefixations = new List<Fixation>();
                    while (i < fixations.Length && fixations[i].Trial == currentTrial)
                    {
                        pagefixations.Add(fixations[i]);
                        i++;
                    }
                    i--;
                    if (FixationsService.IsLeagalPageVistFixations(pagefixations))
                    {
                        this.m_Pages_Sequence.Add(currentPage);
                    }


                }
            }
            public static List<Fixation> FilterFunction(List<Fixation> original,out Dictionary<string, int> dict, NumericExpression numeric)
            {
                dict = new Dictionary<string, int>();
                List<Fixation> fixationsAfterFilter = new List<Fixation>();
                IEnumerable<double> valuesForStandardDevision = original.Select(fix => numeric(fix));
                if (valuesForStandardDevision.Count() == 0)
                    return fixationsAfterFilter;
                List<double> listForTest = valuesForStandardDevision.ToList();
                IEnumerable<double> standardDevisionGrades = StandardDevision.ComputeStandardDevisionGrades(valuesForStandardDevision);
                List<double> listForTest2 = standardDevisionGrades.ToList();

                int length = standardDevisionGrades.Count();
                for (int i = 0; i < length; ++i)
                {
                    Fixation fixation = original[i];
                    if (!dict.ContainsKey(fixation.Trial))
                        dict.Add(fixation.Trial, 0);
                    if (Math.Abs(standardDevisionGrades.ElementAt(i)) > standardDevisionAllowed)
                    {
                            dict[fixation.Trial]++;
                    }
                    else
                        fixationsAfterFilter.Add(fixation);
                }
                return fixationsAfterFilter;
            }
        }
        public static void CreateDatasetFilterTrialText()
        {
            try
            {
                standardDevisionAllowed = double.Parse(ConfigurationService.StandardDeviation);
            }
            catch
            {
                ExcelLogger.ExcelLoggerService.AddLog(new ExcelLogger.Log() { FileName = ConfigurationService.CONFIG_FILE, Description = "Standard Deviation Value Is Not A Number" });
                standardDevisionAllowed = 2.5;
            }
            // grouping the fixations of each participant from all the trials (pages)
            IEnumerable<IGrouping<string, KeyValuePair<string, List<Fixation>>>> fixationsGroupingByParticipant = FixationsService.fixationSetToFixationListDictionary.GroupBy(fl => fl.Key.Split('\t')[0]);

            foreach (IGrouping<string, KeyValuePair<string, List<Fixation>>> participantFixations in fixationsGroupingByParticipant)
            {
                List<Fixation> fixationListPerParticipant = participantFixations.SelectMany(list => list.Value).ToList();
                IEnumerable<IGrouping<string, Fixation>> fixationsGroupingByStimulus = fixationListPerParticipant.GroupBy(fl => fl.Stimulus_Tokens[0]);
                foreach (IGrouping<string, Fixation> participantFixationsInStimulus in fixationsGroupingByStimulus)
                {
                    List<Fixation> fixationListPerParticipantPerStimulus = participantFixationsInStimulus.ToList();
                    if (fixationListPerParticipantPerStimulus.Count > 0)
                    {
                        string s1participant = participantFixations.Key;
                        string s2stimulus = participantFixationsInStimulus.Key;
                        FilteredTrialTextPerParticipant participantFiltered = new FilteredTrialTextPerParticipant(participantFixations.Key, participantFixationsInStimulus.Key, fixationListPerParticipantPerStimulus);
                        filteredTrialTextPerParticipants.Add(participantFiltered);

                    }

                }
            }
        }
        private static void CreateFile(FilteredTrialTextPerParticipant item)
        {
            string participant = item.Participant;
            List<Fixation> table = new List<Fixation>();
            List<Fixation> values = item.m_Fixations_Text;

            List<Saccade> saccadeTable = new List<Saccade>();
            // all fixations
            table.AddRange(values);
            CreateExcelFromStringTable(participant, " All fixations", table, null);
            table.Clear();

            // all fixations after filter
            values = item.All_Fixations_Duration_Filter;
            table.AddRange(values);
            CreateExcelFromStringTable(participant, " All fixations after duration filter", table, null);
            table.Clear();

            // Progressive fixations:

            values = item.m_Progressive_Fixations;
            table.AddRange(values);
            CreateExcelFromStringTable(participant, " All progressive fixations", table, null);
            table.Clear();

            values = item.Progressive_Fixations_Duration_Filter;
            table.AddRange(values);
            CreateExcelFromStringTable(participant, " Progressive fixations after duration filter", table, null);
            table.Clear();

            values = item.m_Progressive_Fixations;
            saccadeTable.AddRange(LastFixationsToSaccades(values, fix => fix.DistanceToPreviousFixation()));
            CreateExcelFromStringTable(participant, " All progressive saccades", saccadeTable, null);
            saccadeTable.Clear();

            values = item.Progressive_Saccade_Length_Filter;
            saccadeTable.AddRange(LastFixationsToSaccades(values, fix => fix.DistanceToPreviousFixation()));
            CreateExcelFromStringTable(participant, " Progressive saccades after filter", saccadeTable, null);
            saccadeTable.Clear();

            values = item.m_Progressive_Fixations;
            saccadeTable.AddRange(LastFixationsToSaccades(values, fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X)));
            CreateExcelFromStringTable(participant, " All progressive saccades length X", saccadeTable, null);
            saccadeTable.Clear();

            values = item.Progressive_Saccade_Length_X_Filter;
            saccadeTable.AddRange(LastFixationsToSaccades(values, fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X)));
            CreateExcelFromStringTable(participant, " Progressive saccades length X after filter", saccadeTable, null);
            saccadeTable.Clear();


            // Regressive Fixations:

            values = item.m_Regressive_Fixations;
            table.AddRange(values);
            CreateExcelFromStringTable(participant, " All Regressive fixations", table, null);
            table.Clear();

            values = item.Regressive_Fixations_Duration_Filter;
            table.AddRange(values);
            CreateExcelFromStringTable(participant, " Regressive fixations after duration filter", table, null);
            table.Clear();

            values = item.m_Regressive_Fixations;
            saccadeTable.AddRange(LastFixationsToSaccades(values, fix => fix.DistanceToPreviousFixation()));
            CreateExcelFromStringTable(participant, " All Regressive saccades", saccadeTable, null);
            saccadeTable.Clear();

            values = item.Regressive_Saccade_Length_Filter;
            saccadeTable.AddRange(LastFixationsToSaccades(values, fix => fix.DistanceToPreviousFixation()));
            CreateExcelFromStringTable(participant, " Regressive saccades after filter", saccadeTable, null);
            saccadeTable.Clear();

            values = item.m_Regressive_Fixations;
            saccadeTable.AddRange(LastFixationsToSaccades(values, fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X)));
            CreateExcelFromStringTable(participant, " All Regressive saccades length X", saccadeTable, null);
            saccadeTable.Clear();

            values = item.Regressive_Saccade_Length_X_Filter;
            saccadeTable.AddRange(LastFixationsToSaccades(values, fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X)));
            CreateExcelFromStringTable(participant, " Regressive saccades length X after filter", saccadeTable, null);
            saccadeTable.Clear();


        }
        /*
         * Creates files only for test the filter on fixations function
         */
        public static void CreateFilesForTest()
        {
            int total = filteredTrialTextPerParticipants.Count();
            int lasts = total % 10;
            int i = 0;
            for(i = 0; i < total - lasts; i+=10)
            {

                FilteredTrialTextPerParticipant item1 = filteredTrialTextPerParticipants[i];
                FilteredTrialTextPerParticipant item2 = filteredTrialTextPerParticipants[i + 1];
                FilteredTrialTextPerParticipant item3 = filteredTrialTextPerParticipants[i + 2];
                FilteredTrialTextPerParticipant item4 = filteredTrialTextPerParticipants[i + 3];
                FilteredTrialTextPerParticipant item5 = filteredTrialTextPerParticipants[i + 4];
                FilteredTrialTextPerParticipant item6 = filteredTrialTextPerParticipants[i + 5];
                FilteredTrialTextPerParticipant item7 = filteredTrialTextPerParticipants[i + 6];
                FilteredTrialTextPerParticipant item8 = filteredTrialTextPerParticipants[i + 7];
                FilteredTrialTextPerParticipant item9 = filteredTrialTextPerParticipants[i + 8];
                FilteredTrialTextPerParticipant item10 = filteredTrialTextPerParticipants[i + 9];

                var t1 = new Thread(() => CreateFile(item1));
                t1.Start();
                var t2 = new Thread(() => CreateFile(item2));
                t2.Start();
                var t3 = new Thread(() => CreateFile(item3));
                t3.Start();
                var t4 = new Thread(() => CreateFile(item4));
                t4.Start();
                var t5 = new Thread(() => CreateFile(item5));
                t5.Start();
                var t6 = new Thread(() => CreateFile(item6));
                t6.Start();
                var t7 = new Thread(() => CreateFile(item7));
                t7.Start();
                var t8 = new Thread(() => CreateFile(item8));
                t8.Start();
                var t9 = new Thread(() => CreateFile(item9));
                t9.Start();
                var t10 = new Thread(() => CreateFile(item10));
                t10.Start();

                t1.Join();
                t2.Join();
                t3.Join();
                t4.Join();
                t5.Join();
                t6.Join();
                t7.Join();
                t8.Join();
                t9.Join();
                t10.Join();
                Console.WriteLine("Completed " + i.ToString() + " filter files out of " + total.ToString());


            }

            for (int j = i; j < total; j++)
            {
                FilteredTrialTextPerParticipant item1 = filteredTrialTextPerParticipants[j];
                CreateFile(item1);
            }
            Console.WriteLine("Completed filter files");

        }
        public class Saccade
        {
            public string Participant { get; set; }
            public string Trial { get; set; }
            public string Stimulus { get; set; }
            [Description("Last Fixation Index")]
            public long Index { get; set; }
            [Description("Saccade Length")]
            public double Saccade_Length { get; set; }
        }
        private static List<Saccade> LastFixationsToSaccades(List<Fixation> fixations, Func<Fixation, double> saccadeFunction)
        {
            List<Saccade> saccades = new List<Saccade>();
            fixations.ForEach(fix => saccades.Add( new Saccade
            {
                Participant = fix.Participant,
                Stimulus = fix.Stimulus,
                Index = fix.Index,
                Trial = fix.Trial,
                Saccade_Length = saccadeFunction(fix)
            }));
            return saccades;
        }


        public static void CreateExcelFromStringTable<T>(string participant ,string fileName, IEnumerable<T> table, Func<ExcelWorksheet, int> editExcelFunc)
        {
            using (var wb = new ExcelPackage())
            {
                ExcelWorksheet ws = wb.Workbook.Worksheets.Add("Inserting Tables");
/*
                string islogs = "Logs";
                string isAOIFiltered = "AOI - Filtered";
                string isPageFiltered = "Page - Filtered";
                string isTextFiltered = "Text - Filtered";
                if (!fileName.Contains(islogs))
                {
                    ws.View.FreezePanes(2, 4);
                }
                if (fileName.Contains(isAOIFiltered))
                {
                    ws.View.FreezePanes(2, 6);
                }
                if (fileName.Contains(ConfigurationService.FourthExcelFileName))
                {
                    ws.View.FreezePanes(2, 3);
                }
*/
                ExcelRangeBase range = ws.Cells[1, 1].LoadFromCollectionFiltered(table);

                ws.Cells[ws.Dimension.Address].AutoFitColumns();
                editExcelFunc?.Invoke(ws);
                DialogResult dialogResult; // = DialogResult.Retry;
                do
                {
                    try
                    {
                        string path = FixationsService.outputPath + "/" + FixationsService.outputTextString + " - Filters/Test Filters/" + participant;
                        if (!Directory.Exists(path))
                        {
                            DirectoryInfo di = Directory.CreateDirectory(path);
                        }
                        wb.SaveAs(new FileInfo(path + "/" + fileName + ConfigurationService.ExcelFilesExtension));
                    
                        dialogResult = DialogResult.Abort;
                    }
                    catch (InvalidOperationException e)
                    {
                        Console.Write(e.InnerException.InnerException.Message);
                        string errorDescription = "";
                        errorDescription += e.InnerException?.InnerException?.Message + Environment.NewLine;
                        errorDescription += "Check If The File We Trying to overwrite is already open!!!";
                        dialogResult = MessageBox.Show(errorDescription, "Error In Saving File " + fileName, MessageBoxButtons.RetryCancel, MessageBoxIcon.Error);
                    }
                } while (dialogResult == DialogResult.Retry);
            }
        }
    }
}
