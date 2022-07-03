using EM_Analyzer.ModelClasses;
using EM_Analyzer.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EM_Analyzer.ExcelsFilesMakers
{
    class FourthFileAfterProccessing
    {
        public static void MakeExcelFile()
        {
            // grouping the fixations of each participant from all the trials (pages)
            IEnumerable<IGrouping<string, KeyValuePair<string, List<Fixation>>>> fixationsGroupingByParticipant = FixationsService.fixationSetToFixationListDictionary.GroupBy(fl => fl.Key.Split('\t')[0]);
            List<ParticipantText> table = new List<ParticipantText>();

            foreach (IGrouping<string, KeyValuePair<string, List<Fixation>>> participantFixations in fixationsGroupingByParticipant)
            {
                List<Fixation> fixationListPerParticipant = participantFixations.SelectMany(list => list.Value).ToList();

                IEnumerable<IGrouping<string, Fixation>> fixationsGroupingByStimulus = fixationListPerParticipant.GroupBy(fl => fl.Stimulus_Tokens[0]);

                foreach (IGrouping<string, Fixation> participantFixationsInStimulus in fixationsGroupingByStimulus)
                {
                    List<Fixation> fixationListPerParticipantPerStimulus = participantFixationsInStimulus.ToList();
                    if (fixationListPerParticipantPerStimulus.Count > 0)
                        table.Add(new ParticipantText(participantFixationsInStimulus.Key, participantFixations.Key, fixationListPerParticipantPerStimulus));

                }

            }
            ExcelsService.CreateExcelFromStringTable(ConfigurationService.FourthExcelFileName, table, null);
        }
        private class ParticipantText
        {
            [Description("Participant")]
            public string Participant { get; set; }
            [Description("Stimulus")]
            public string Stimulus { get; set; }

            private List<Fixation> m_fixations_text;

            private double m_Mean_Fixation_Duration;
            [Description("Mean Fixation Duration")]
            public double Mean_Fixation_Duration
            {
                get
                {
                    if (this.m_Mean_Fixation_Duration == -1)
                    {
                        double duration_sum = this.m_fixations_text.Sum(fix => fix.Event_Duration);
                        this.m_Mean_Fixation_Duration = duration_sum / this.m_fixations_text.Count;
                    }
                    return this.m_Mean_Fixation_Duration;
                }
            }
            private double m_SD_Fixation_Duration;
            [Description("SD Fixation Duration")]
            public double SD_Fixation_Duration
            {
                get
                {
                    if (this.m_SD_Fixation_Duration == -1)
                    {
                        IEnumerable<double> durations = this.m_fixations_text.Select(fix => fix.Event_Duration).ToList();
                        this.m_SD_Fixation_Duration = StandardDevision.ComputeStandardDevision(durations);
                    }
                    return this.m_SD_Fixation_Duration;
                }
            }



            private int m_Total_Fixation_Number;
            [Description("Total Fixation Number")]
            public int Total_Fixation_Number
            {
                get
                {
                    if (this.m_Total_Fixation_Number == -1)
                    {
                        this.m_Total_Fixation_Number = this.m_fixations_text.Count;
                    }
                    return this.m_Total_Fixation_Number;
                }
            }


            private double m_Mean_Progressive_Fixation_Duration;
            [Description("Mean Progressive Fixation Duration")]
            public double Mean_Progressive_Fixation_Duration
            {
                get
                {
                    if (this.m_Mean_Progressive_Fixation_Duration == -1)
                    {
                        double duration_sum = this.Progressive_Fixations.Sum(fix => fix.Event_Duration);
                        this.m_Mean_Progressive_Fixation_Duration = duration_sum / this.Progressive_Fixations.Count;
                    }
                    return this.m_Mean_Progressive_Fixation_Duration;
                }
            }
            private double m_SD_Progressive_Fixation_Duration;
            [Description("SD Progressive Fixation Duration")]
            public double SD_Progressive_Fixation_Duration
            {
                get
                {
                    if (this.m_SD_Progressive_Fixation_Duration == -1)
                    {
                        IEnumerable<double> durations = this.Progressive_Fixations.Select(fix => fix.Event_Duration).ToList();
                        this.m_SD_Progressive_Fixation_Duration = StandardDevision.ComputeStandardDevision(durations);
                    }
                    return this.m_SD_Progressive_Fixation_Duration;
                }
            }
            [Description("Progressive Fixation Number")]
            public int Progressive_Fixation_Number
            {
                get
                {
                    return this.Progressive_Fixations.Count;
                }
            }

            private double m_Mean_Progressive_Saccade_Length;
            [Description("Mean Progressive Saccade Length")]
            public double Mean_Progressive_Saccade_Length
            {
                get
                {
                    if (this.m_Mean_Progressive_Saccade_Length == -1)
                    {
                        this.m_Mean_Progressive_Saccade_Length = this.Progressive_Fixations.Sum(fix => fix.DistanceToPreviousFixation()) / this.Progressive_Fixations.Count;
                    }
                    return this.m_Mean_Progressive_Saccade_Length;
                }
            }
            private double m_SD_Progressive_Saccade_Length;
            [Description("SD Progressive Saccade Length")]
            public double SD_Progressive_Saccade_Length
            {
                get
                {
                    if (this.m_SD_Progressive_Saccade_Length == -1)
                    {
                        IEnumerable<double> distances = this.Progressive_Fixations.Select(fix => fix.DistanceToPreviousFixation()).ToList();
                        this.m_SD_Progressive_Saccade_Length = StandardDevision.ComputeStandardDevision(distances);
                    }
                    return this.m_SD_Progressive_Saccade_Length;
                }
            }

            private double m_Mean_Progressive_Saccade_Length_X;
            [Description("Mean Progressive Saccade Length X")]
            public double Mean_Progressive_Saccade_Length_X
            {
                get
                {
                    if (this.m_Mean_Progressive_Saccade_Length_X == -1)
                    {
                        this.m_Mean_Progressive_Saccade_Length_X = this.Progressive_Fixations.Sum(fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X)) / this.Progressive_Fixations.Count;
                    }
                    return this.m_Mean_Progressive_Saccade_Length_X;
                }
            }
            private double m_SD_Progressive_Saccade_Length_X;
            [Description("SD Progressive Saccade Length X")]
            public double SD_Progressive_Saccade_Length_X
            {
                get
                {
                    if (this.m_SD_Progressive_Saccade_Length_X == -1)
                    {
                        IEnumerable<double> distances = this.Progressive_Fixations.Select(fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X)).ToList();
                        this.m_SD_Progressive_Saccade_Length_X = StandardDevision.ComputeStandardDevision(distances);
                    }
                    return this.m_SD_Progressive_Saccade_Length_X;
                }
            }


            private double m_Mean_Regressive_Fixation_Duration;
            [Description("Mean Regressive Fixation Duration")]
            public double Mean_Regressive_Fixation_Duration
            {
                get
                {
                    if (this.m_Mean_Regressive_Fixation_Duration == -1)
                    {
                        double duration_sum = this.Regressive_Fixations.Sum(fix => fix.Event_Duration);
                        this.m_Mean_Regressive_Fixation_Duration = duration_sum / this.Regressive_Fixations.Count;
                    }
                    return this.m_Mean_Regressive_Fixation_Duration;
                }
            }
            private double m_SD_Regressive_Fixation_Duration;
            [Description("SD Regressive Fixation Duration")]
            public double SD_Regressive_Fixation_Duration
            {
                get
                {
                    if (this.m_SD_Regressive_Fixation_Duration == -1)
                    {
                        IEnumerable<double> durations = this.Regressive_Fixations.Select(fix => fix.Event_Duration).ToList();
                        this.m_SD_Regressive_Fixation_Duration = StandardDevision.ComputeStandardDevision(durations);
                    }
                    return this.m_SD_Regressive_Fixation_Duration;
                }
            }
            [Description("Regressive Fixation Number")]
            public int Regressive_Fixation_Number
            {
                get
                {
                    return this.Regressive_Fixations.Count;
                }
            }

            private double m_Mean_Regressive_Saccade_Length;
            [Description("Mean Regressive Saccade Length")]
            public double Mean_Regressive_Saccade_Length
            {
                get
                {
                    if (this.m_Mean_Regressive_Saccade_Length == -1)
                    {
                        this.m_Mean_Regressive_Saccade_Length = this.Regressive_Fixations.Sum(fix => fix.DistanceToPreviousFixation()) / this.Regressive_Fixations.Count;
                    }
                    return this.m_Mean_Regressive_Saccade_Length;
                }
            }
            private double m_SD_Regressive_Saccade_Length;
            [Description("SD Regressive Saccade Length")]
            public double SD_Regressive_Saccade_Length
            {
                get
                {
                    if (this.m_SD_Regressive_Saccade_Length == -1)
                    {
                        IEnumerable<double> distances = this.Regressive_Fixations.Select(fix => fix.DistanceToPreviousFixation()).ToList();
                        this.m_SD_Regressive_Saccade_Length = StandardDevision.ComputeStandardDevision(distances);
                    }
                    return this.m_SD_Regressive_Saccade_Length;
                }
            }

            private double m_Mean_Regressive_Saccade_Length_X;
            [Description("Mean Regressive Saccade Length X")]
            public double Mean_Regressive_Saccade_Length_X
            {
                get
                {
                    if (this.m_Mean_Regressive_Saccade_Length_X == -1)
                    {
                        this.m_Mean_Regressive_Saccade_Length_X = this.Regressive_Fixations.Sum(fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X)) / this.Regressive_Fixations.Count;
                    }
                    return this.m_Mean_Regressive_Saccade_Length_X;
                }
            }
            private double m_SD_Regressive_Saccade_Length_X;
            [Description("SD Regressive Saccade Length X")]
            public double SD_Regressive_Saccade_Length_X
            {
                get
                {
                    if (this.m_SD_Regressive_Saccade_Length_X == -1)
                    {
                        IEnumerable<double> distances = this.Regressive_Fixations.Select(fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X)).ToList();
                        this.m_SD_Regressive_Saccade_Length_X = StandardDevision.ComputeStandardDevision(distances);
                    }
                    return this.m_SD_Regressive_Saccade_Length_X;
                }
            }
            private double m_Pupil_Diameter;

            [Description("Mean Pupil Diameter [mm]")]
            public double Mean_Pupil_Diameter
            {
                get
                {
                    if (this.m_Pupil_Diameter == -1)
                    {
                        double sum = this.m_fixations_text.Sum(fix => fix.Fixation_Average_Pupil_Diameter);
                        this.m_Pupil_Diameter = sum / this.m_fixations_text.Count;
                    }
                    return this.m_Pupil_Diameter;
                }
            }

            private List<int> m_Pages_Sequence;

            [Description("Page Moves")]
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
            [Description("Page Regressions")]
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
            [Description("Page Visits")]
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



            [EpplusIgnore]
            public List<Fixation> All_Fixations { get; set; }
            private List<Fixation> m_Progressive_Fixations;
            private List<Fixation> Progressive_Fixations
            {
                get
                {
                    if (this.m_Progressive_Fixations == null)
                    {
                        this.InitializeProgressiveAndRegressiveFixationsList();
                    }

                    return this.m_Progressive_Fixations;
                }
            }

            private List<Fixation> m_Regressive_Fixations;
            private List<Fixation> Regressive_Fixations
            {
                get
                {
                    if (this.m_Regressive_Fixations == null)
                    {
                        this.InitializeProgressiveAndRegressiveFixationsList();
                    }
                    return this.m_Regressive_Fixations;
                }
            }




            public ParticipantText(string Stimulus, string Participant, List<Fixation> fixations)
            {

                this.Stimulus = Stimulus;
                this.Participant = Participant;
                
                // All fixation, include first fixation in every trial is Only for the calculation 
                // of the other fixations as progressive and regressive
                this.All_Fixations = fixations;
                
                // all fixations of text without the first one in every trial (and all the aoi's > 0)
                this.m_fixations_text = new List<Fixation>();
                List<IGrouping<string, Fixation>> groupingTrials = All_Fixations.GroupBy(x => x.Trial).ToList();
                foreach (IGrouping<string, Fixation>  item in groupingTrials)
                    m_fixations_text.AddRange(item.Skip(1).ToList());
                this.m_Total_Fixation_Number = -1;
                this.m_Mean_Fixation_Duration = -1;
                this.m_SD_Fixation_Duration = -1;
                this.m_Mean_Progressive_Fixation_Duration = -1;
                this.m_SD_Progressive_Fixation_Duration = -1;
                this.m_Mean_Progressive_Saccade_Length = -1;
                this.m_SD_Progressive_Saccade_Length = -1;
                this.m_Mean_Progressive_Saccade_Length_X = -1;
                this.m_SD_Progressive_Saccade_Length_X = -1;
                this.m_Progressive_Fixations = null;
                this.m_Mean_Regressive_Fixation_Duration = -1;
                this.m_SD_Regressive_Fixation_Duration = -1;
                this.m_Mean_Regressive_Saccade_Length = -1;
                this.m_SD_Regressive_Saccade_Length = -1;
                this.m_Mean_Regressive_Saccade_Length_X = -1;
                this.m_SD_Regressive_Saccade_Length_X = -1;
                this.m_Pupil_Diameter = -1;
                this.m_Regressive_Fixations = null;
                this.m_Pages_Sequence = null;




            }
            private void InitializePagesSequence()
            {
                this.m_Pages_Sequence = new List<int>();
                Fixation[] fixations = m_fixations_text.ToArray();
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
                if (this.m_Pages_Sequence.Count == 0)
                {
                    this.m_Pages_Sequence.Add(0);
                }
                /*
                List<int> copy = new List<int>();
                int prev = this.m_Pages_Sequence[0];
                copy.Add(prev);
                for (int i = 1; i < this.m_Pages_Sequence.Count; i++)
                {
                    int curr = this.m_Pages_Sequence[i];
                    if (prev != curr)
                    {
                        copy.Add(curr);
                    }
                    prev = curr;


                }
                this.m_Pages_Sequence = copy;
                */
            }
            private void InitializeProgressiveAndRegressiveFixationsList()
            {
                this.m_Progressive_Fixations = new List<Fixation>();
                this.m_Regressive_Fixations = new List<Fixation>();
                Fixation[] fixations = this.All_Fixations.ToArray();
                int currentPage = fixations[0].Page;
                for (int i = 1; i < fixations.Length; ++i)
                {
                    if (currentPage != fixations[i].Page)
                    {
                        currentPage = fixations[i].Page;
                        continue;
                    }

                    fixations[i].Previous_Fixation = fixations[i - 1];
                    if (fixations[i - 1].IsBeforeThan(fixations[i]))
                        this.m_Progressive_Fixations.Add(fixations[i]);
                    else
                        this.m_Regressive_Fixations.Add(fixations[i]);
                }
            }
        }

    }
}
