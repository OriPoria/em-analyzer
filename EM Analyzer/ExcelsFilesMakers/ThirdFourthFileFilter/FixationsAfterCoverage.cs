using EM_Analyzer.ModelClasses;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static EM_Analyzer.ExcelsFilesMakers.ThirdFileConsideringCoverage;

namespace EM_Analyzer.ExcelsFilesMakers.ThirdFourFilter
{
    class ThirdFourthFilter
    {
        public static double standardDevisionAllowed;
        public delegate double NumericExpression(Fixation value);
        public static List<FilteredTrialTextPerParticipant> filteredTrialTextPerParticipants = new List<FilteredTrialTextPerParticipant>();

        public class FilteredTrialTextPerParticipant
        {
            public string Participant { get; set; }
            public string Stimulus { get; set; }
            public HashSet<string> All_Trials = new HashSet<string>();
            // all fixations
            private List<Fixation> allFixations;
            private List<Fixation> m_Fixations_Text;
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
            private List<Fixation> m_Progressive_Fixations;
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
            private List<Fixation> m_Regressive_Fixations;
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


            public FilteredTrialTextPerParticipant (string Stimulus, string Participant, List<Fixation> fixations)
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
                standardDevisionAllowed = 1.0;
            }
            catch
            {
                ExcelLogger.ExcelLoggerService.AddLog(new ExcelLogger.Log() { FileName = ConfigurationService.CONFIG_FILE, Description = "Standard Deviation Value Is Not A Number" });
                standardDevisionAllowed = 2.5;
            }
            // grouping the fixations of each participant from all the trials (pages)
            IEnumerable<IGrouping<string, KeyValuePair<string, List<Fixation>>>> fixationsGroupingByParticipant = FixationsService.fixationSetToFixationListDictionary.GroupBy(fl => fl.Value[0].Participant);

            foreach (IGrouping<string, KeyValuePair<string, List<Fixation>>> participantFixations in fixationsGroupingByParticipant)
            {
                List<Fixation> fixationList = participantFixations.SelectMany(list => list.Value).ToList();
                FilteredTrialTextPerParticipant participantFiltered = new FilteredTrialTextPerParticipant(participantFixations.Key, fixationList[0].Stimulus_Tokens[0], fixationList);
                filteredTrialTextPerParticipants.Add(participantFiltered);
            }
        }
    }
}
