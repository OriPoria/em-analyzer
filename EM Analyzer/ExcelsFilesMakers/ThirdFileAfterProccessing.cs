using EM_Analyzer.Interfaces;
using EM_Analyzer.ModelClasses;
using EM_Analyzer.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;

namespace EM_Analyzer.ExcelsFilesMakers
{
    class ThirdFileAfterProccessing
    {
        public static void MakeExcelFile()
        {
            List<Fixation>[] fixationsLists = FixationsService.fixationSetToFixationListDictionary.Values.ToArray();
            List<ParticipantTrial> table = new List<ParticipantTrial>();
            
            foreach(List<Fixation> fixations in fixationsLists)
            {
                table.Add(new ParticipantTrial(fixations[0].Trial, fixations[0].Stimulus, fixations[0].Participant));
            }
            ExcelsService.CreateExcelFromStringTable(ConfigurationService.ThirdExcelFileName, table, null);
        }


        private class ParticipantTrial : ITrialClassForConsideringCoverage
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
                        double duration_sum = this.Fixations.Sum(fix => fix.Event_Duration);
                        this.m_Mean_Fixation_Duration = duration_sum / this.Fixations.Count;
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
                        IEnumerable<double> durations = this.Fixations.Select(fix => fix.Event_Duration).ToList();
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
                        this.m_Total_Fixation_Number = this.Fixations.Count;
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
                        double sum = this.Fixations.Sum(fix => fix.Fixation_Average_Pupil_Diameter);
                        this.m_Pupil_Diameter = sum / this.Fixations.Count;
                    }
                    return this.m_Pupil_Diameter;
                }
            }



            [EpplusIgnore]
            public List<Fixation> Fixations { get; set; }

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


            public ParticipantTrial(string Trial, string Stimulus, string Participant)
            {
                this.Trial = Trial;
                this.Stimulus = Stimulus;
                this.Participant = Participant;

                /*
                 Critical Point: Removes all the fixations with no String AOI (after dealing with exceptions) !
                 */
                this.Fixations = FixationsService.fixationSetToFixationListDictionary[this.Participant + '\t' + this.Trial + '\t' + this.Stimulus];
                this.Fixations.RemoveAll(fix => fix.IsStringAOI);

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

            }
            

            private void InitializeProgressiveAndRegressiveFixationsList()
            {
                this.m_Progressive_Fixations = new List<Fixation>();
                this.m_Regressive_Fixations = new List<Fixation>();
                Fixation[] fixations = this.Fixations.ToArray();

                for (int i = 1; i < fixations.Length; ++i)
                {
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
