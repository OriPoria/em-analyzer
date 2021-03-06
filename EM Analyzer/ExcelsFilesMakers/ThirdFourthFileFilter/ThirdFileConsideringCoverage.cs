using EM_Analyzer.ExcelsFilesMakers.ThirdFourFilter;
using EM_Analyzer.ModelClasses;
using EM_Analyzer.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static EM_Analyzer.ExcelsFilesMakers.ThirdFourFilter.ThirdFourthFilter;

namespace EM_Analyzer.ExcelsFilesMakers
{
    class ThirdFileConsideringCoverage
    {
        public static void MakeExcelFile()
        {
            List<FilteredTrialTextPerParticipant> perParticipants = filteredTrialTextPerParticipants;
            List<ParticipantTrialAfterFilter> table = new List<ParticipantTrialAfterFilter>();

            foreach (FilteredTrialTextPerParticipant item in perParticipants)
            {
                foreach (string trial in item.All_Trials)
                {
                    table.Add(new ParticipantTrialAfterFilter(trial, item.All_Fixations_Duration_Filter[0].Stimulus,
                        item.All_Fixations_Duration_Filter[0].Participant, item));
                }
                ExcelsService.CreateExcelFromStringTable("Page - Filtered By Participant", table, null);

            }
        }


        public class ParticipantTrialAfterFilter
        {
            [Description("Participant")]
            public string Participant { get; set; }
            [Description("Trial")]
            public string Trial { get; set; }
            [Description("Stimulus")]
            public string Stimulus { get; set; }

            private double m_Mean_Fixation_Duration;
            [Description("Mean Fixation Duration")]
            public double Mean_Fixation_Duration
            {
                get
                {
                    if (this.m_Mean_Fixation_Duration == -1)
                    {
                        double duration_sum = this.m_Fixations_Duration_Filtered.Sum(fix => fix.Event_Duration);
                        this.m_Mean_Fixation_Duration = duration_sum / this.m_Fixations_Duration_Filtered.Count;
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
                        IEnumerable<double> durations = this.m_Fixations_Duration_Filtered.Select(fix => fix.Event_Duration).ToList();
                        this.m_SD_Fixation_Duration = StandardDevision.ComputeStandardDevision(durations);
                    }
                    return this.m_SD_Fixation_Duration;
                }
            }

            private List<Fixation> m_Total_Fixation_Page;

            private int m_Total_Fixation_Number;
            [Description("Total Fixation Number")]
            public int Total_Fixation_Number
            {
                get
                {
                    if (this.m_Total_Fixation_Number == -1)
                    {
                        this.m_Total_Fixation_Number = this.m_Total_Fixation_Page.Count;
                    }
                    return this.m_Total_Fixation_Number;
                }
            }
            [Description("Total Fixation Filtered")]
            public int Total_Fixation_Filtered
            {
                get
                {
                    if (!filteredTrialTextPerParticipant.Trial_Elimination_Fixations_Duration.ContainsKey(Trial)) return 0;
                    return filteredTrialTextPerParticipant.Trial_Elimination_Fixations_Duration[Trial];
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
                        double duration_sum = this.m_Progressive_Fixations_Duration_Filtered.Sum(fix => fix.Event_Duration);
                        this.m_Mean_Progressive_Fixation_Duration = duration_sum / this.m_Progressive_Fixations_Duration_Filtered.Count;
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
                        IEnumerable<double> durations = this.m_Progressive_Fixations_Duration_Filtered.Select(fix => fix.Event_Duration).ToList();
                        this.m_SD_Progressive_Fixation_Duration = StandardDevision.ComputeStandardDevision(durations);
                    }
                    return this.m_SD_Progressive_Fixation_Duration;
                }
            }
            private List<Fixation> m_Progressive_Fixations;
            [Description("Progressive Fixation Number")]
            public int Progressive_Fixation_Number
            {
                get
                {
                    return this.m_Progressive_Fixations.Count;
                }
            }
            [Description("Progressive Fixation Filtered")]
            public int Progressive_Fixation_Filtered
            {
                get
                {
                    if (!filteredTrialTextPerParticipant.Trial_Elimination_Progressive_Duration.ContainsKey(Trial)) return 0;
                    return filteredTrialTextPerParticipant.Trial_Elimination_Progressive_Duration[Trial];
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
                        this.m_Mean_Progressive_Saccade_Length = this.m_Progressive_Fixations_Saccade_Length_Filtered.Sum(fix => fix.DistanceToPreviousFixation()) / this.m_Progressive_Fixations_Saccade_Length_Filtered.Count;
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
                        IEnumerable<double> distances = this.m_Progressive_Fixations_Saccade_Length_Filtered.Select(fix => fix.DistanceToPreviousFixation()).ToList();
                        this.m_SD_Progressive_Saccade_Length = StandardDevision.ComputeStandardDevision(distances);
                    }
                    return this.m_SD_Progressive_Saccade_Length;
                }
            }
            [Description("Progressive Saccade Length Filtered")]
            public double Progressive_Saccade_Length_Filtered
            {
                get
                {
                    if (!filteredTrialTextPerParticipant.Trial_Elimination_Prog_Saccade_Length.ContainsKey(Trial)) return 0;
                    return filteredTrialTextPerParticipant.Trial_Elimination_Prog_Saccade_Length[Trial];
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
                        this.m_Mean_Progressive_Saccade_Length_X = this.m_Progressive_Fixations_Saccade_Length_X_Filtered.Sum(fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X)) / this.m_Progressive_Fixations_Saccade_Length_X_Filtered.Count;
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
                        IEnumerable<double> distances = this.m_Progressive_Fixations_Saccade_Length_X_Filtered.Select(fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X)).ToList();
                        this.m_SD_Progressive_Saccade_Length_X = StandardDevision.ComputeStandardDevision(distances);
                    }
                    return this.m_SD_Progressive_Saccade_Length_X;
                }
            }
            [Description("Progressive Saccade Length X Filtered")]
            public double Progressive_Saccade_Length_X_Filtered
            {
                get
                {
                    if (!filteredTrialTextPerParticipant.Trial_Elimination_Prog_Saccade_X_Length.ContainsKey(Trial)) return 0;
                    return filteredTrialTextPerParticipant.Trial_Elimination_Prog_Saccade_X_Length[Trial];
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
                        double duration_sum = this.m_Regressive_Fixations_Filtered.Sum(fix => fix.Event_Duration);
                        this.m_Mean_Regressive_Fixation_Duration = duration_sum / this.m_Regressive_Fixations_Filtered.Count;
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
                        IEnumerable<double> durations = this.m_Regressive_Fixations_Filtered.Select(fix => fix.Event_Duration).ToList();
                        this.m_SD_Regressive_Fixation_Duration = StandardDevision.ComputeStandardDevision(durations);
                    }
                    return this.m_SD_Regressive_Fixation_Duration;
                }
            }
            private List<Fixation> m_Regressive_Fixations;
            [Description("Regressive Fixation Number")]
            public int Regressive_Fixation_Number
            {
                get
                {
                    return this.m_Regressive_Fixations.Count;
                }
            }
            [Description("Regressive Fixation Filtered")]
            public int Regressive_Fixation_Duration_Filtered
            {
                get
                {
                    if (!filteredTrialTextPerParticipant.Trial_Elimination_Regressive_Duration.ContainsKey(Trial)) return 0;
                    return filteredTrialTextPerParticipant.Trial_Elimination_Regressive_Duration[Trial];
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
                        this.m_Mean_Regressive_Saccade_Length = this.m_Regressive_Fixations_Saccade_Length_Filtered.Sum(fix => fix.DistanceToPreviousFixation()) / this.m_Regressive_Fixations_Saccade_Length_Filtered.Count;
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
                        IEnumerable<double> distances = this.m_Regressive_Fixations_Saccade_Length_Filtered.Select(fix => fix.DistanceToPreviousFixation()).ToList();
                        this.m_SD_Regressive_Saccade_Length = StandardDevision.ComputeStandardDevision(distances);
                    }
                    return this.m_SD_Regressive_Saccade_Length;
                }
            }
            [Description("Regressive Saccade Length Filtered")]
            public double Regressive_Saccade_Length_Filtered
            {
                get
                {
                    if (!filteredTrialTextPerParticipant.Trial_Elimination_Reg_Saccade_Length.ContainsKey(Trial)) return 0;
                    return filteredTrialTextPerParticipant.Trial_Elimination_Reg_Saccade_Length[Trial];
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
                        this.m_Mean_Regressive_Saccade_Length_X = this.m_Regressive_Fixations_Saccade_Length_X_Filtered.Sum(fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X)) / this.m_Regressive_Fixations_Saccade_Length_X_Filtered.Count;
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
                        IEnumerable<double> distances = this.m_Regressive_Fixations_Saccade_Length_X_Filtered.Select(fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X)).ToList();
                        this.m_SD_Regressive_Saccade_Length_X = StandardDevision.ComputeStandardDevision(distances);
                    }
                    return this.m_SD_Regressive_Saccade_Length_X;
                }
            }
            [Description("Regressive Saccade Length X Filtered")]
            public double Regressive_Saccade_Length_X_Filtered
            {
                get
                {
                    if (!filteredTrialTextPerParticipant.Trial_Elimination_Reg_Saccade_X_Length.ContainsKey(Trial)) return 0;
                    return filteredTrialTextPerParticipant.Trial_Elimination_Reg_Saccade_X_Length[Trial];
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
                        double sum = this.m_Fixations_Average_Pupil_Filtered.Sum(fix => fix.Fixation_Average_Pupil_Diameter);
                        this.m_Pupil_Diameter = sum / this.m_Fixations_Average_Pupil_Filtered.Count;
                    }
                    return this.m_Pupil_Diameter;
                }
            }
            [Description("Pupil Diameter Filtered")]
            public double Pupil_Diameter_Filtered
            {
                get
                {
                    if (!filteredTrialTextPerParticipant.Trial_Elimination_Avg_Pupil_Diameter.ContainsKey(Trial)) return 0;
                    return filteredTrialTextPerParticipant.Trial_Elimination_Avg_Pupil_Diameter[Trial];
                }
            }



            private List<Fixation> m_Fixations_Duration_Filtered;

            private List<Fixation> m_Progressive_Fixations_Duration_Filtered;
            private List<Fixation> m_Progressive_Fixations_Saccade_Length_Filtered;
            private List<Fixation> m_Progressive_Fixations_Saccade_Length_X_Filtered;

            private List<Fixation> m_Regressive_Fixations_Filtered;
            private List<Fixation> m_Regressive_Fixations_Saccade_Length_Filtered;
            private List<Fixation> m_Regressive_Fixations_Saccade_Length_X_Filtered;

            private List<Fixation> m_Fixations_Average_Pupil_Filtered;

            private FilteredTrialTextPerParticipant filteredTrialTextPerParticipant;

            public ParticipantTrialAfterFilter(string Trial, string Stimulus, string Participant,
                FilteredTrialTextPerParticipant filteredTrialTextPerParticipant)
            {
                this.Trial = Trial;
                this.Stimulus = Stimulus;
                this.Participant = Participant;
                this.filteredTrialTextPerParticipant = filteredTrialTextPerParticipant;

                this.m_Total_Fixation_Page = filteredTrialTextPerParticipant.m_Fixations_Text.Where(fix => fix.Trial == Trial).ToList();
                this.m_Fixations_Duration_Filtered = filteredTrialTextPerParticipant.All_Fixations_Duration_Filter.Where(fix => fix.Trial == Trial).ToList();

                this.m_Progressive_Fixations = filteredTrialTextPerParticipant.m_Progressive_Fixations.Where(fix => fix.Trial == Trial).ToList();
                this.m_Progressive_Fixations_Duration_Filtered = filteredTrialTextPerParticipant.Progressive_Fixations_Duration_Filter.Where(fix => fix.Trial == Trial).ToList();
                this.m_Progressive_Fixations_Saccade_Length_Filtered = filteredTrialTextPerParticipant.Progressive_Saccade_Length_Filter.Where(fix => fix.Trial == Trial).ToList();
                this.m_Progressive_Fixations_Saccade_Length_X_Filtered = filteredTrialTextPerParticipant.Progressive_Saccade_Length_X_Filter.Where(fix => fix.Trial == Trial).ToList();

                this.m_Regressive_Fixations = filteredTrialTextPerParticipant.m_Regressive_Fixations.Where(fix => fix.Trial == Trial).ToList();
                this.m_Regressive_Fixations_Filtered = filteredTrialTextPerParticipant.Regressive_Fixations_Duration_Filter.Where(fix => fix.Trial == Trial).ToList();
                this.m_Regressive_Fixations_Saccade_Length_Filtered = filteredTrialTextPerParticipant.Regressive_Saccade_Length_Filter.Where(fix => fix.Trial == Trial).ToList();
                this.m_Regressive_Fixations_Saccade_Length_X_Filtered = filteredTrialTextPerParticipant.Regressive_Saccade_Length_X_Filter.Where(fix => fix.Trial == Trial).ToList();

                this.m_Fixations_Average_Pupil_Filtered = filteredTrialTextPerParticipant.Fixations_Average_Pupil_Diameter_Filter.Where(fix => fix.Trial == Trial).ToList();

                this.m_Total_Fixation_Number = -1;
                this.m_Mean_Fixation_Duration = -1;
                this.m_SD_Fixation_Duration = -1;

                this.m_Mean_Progressive_Fixation_Duration = -1;
                this.m_SD_Progressive_Fixation_Duration = -1;

                this.m_Mean_Progressive_Saccade_Length = -1;
                this.m_SD_Progressive_Saccade_Length = -1;

                this.m_Mean_Progressive_Saccade_Length_X = -1;
                this.m_SD_Progressive_Saccade_Length_X = -1;

                this.m_Mean_Regressive_Fixation_Duration = -1;
                this.m_SD_Regressive_Fixation_Duration = -1;

                this.m_Mean_Regressive_Saccade_Length = -1;
                this.m_SD_Regressive_Saccade_Length = -1;

                this.m_Mean_Regressive_Saccade_Length_X = -1;
                this.m_SD_Regressive_Saccade_Length_X = -1;

                this.m_Pupil_Diameter = -1;

            }




        }
    }
}
