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
            ExcelsService.CreateExcelFromStringTable(ConfigurationService.ThirdExcelFileName, table);
        }


        private class ParticipantTrial
        {
            ////[XLColumn(Ignore = true)]
            //public static Dictionary<string, ParticipantTrial> instancesDictionary = new Dictionary<string, ParticipantTrial>();
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


            private double m_Progressive_Fixation_Duration;
            [Description("Progressive Fixation Duration")]
            public double Progressive_Fixation_Duration
            {
                get
                {
                    if (this.m_Progressive_Fixation_Duration == -1)
                    {
                        double duration_sum = this.Progressive_Fixations.Sum(fix => fix.Event_Duration);
                        this.m_Progressive_Fixation_Duration = duration_sum / this.Progressive_Fixations.Count;
                    }
                    return this.m_Progressive_Fixation_Duration;
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

            private double m_Progressive_Saccade_Length;
            [Description("Progressive Saccade Length")]
            public double Progressive_Saccade_Length
            {
                get
                {
                    if (this.m_Progressive_Saccade_Length == -1)
                    {
                        this.m_Progressive_Saccade_Length = this.Progressive_Fixations.Sum(fix => fix.DistanceToPreviousFixation()) / this.Progressive_Fixations.Count;
                    }
                    return this.m_Progressive_Saccade_Length;
                }
            }

            private double m_Progressive_Saccade_Length_X;
            [Description("Progressive Saccade Length X")]
            public double Progressive_Saccade_Length_X
            {
                get
                {
                    if (this.m_Progressive_Saccade_Length_X == -1)
                    {
                        this.m_Progressive_Saccade_Length_X = this.Progressive_Fixations.Sum(fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X)) / this.Progressive_Fixations.Count;
                    }
                    return this.m_Progressive_Saccade_Length_X;
                }
            }


            private double m_Regressive_Fixation_Duration;
            [Description("Regressive Fixation Duration")]
            public double Regressive_Fixation_Duration
            {
                get
                {
                    if (this.m_Regressive_Fixation_Duration == -1)
                    {
                        double duration_sum = this.Regressive_Fixations.Sum(fix => fix.Event_Duration);
                        this.m_Regressive_Fixation_Duration = duration_sum / this.Regressive_Fixations.Count;
                    }
                    return this.m_Regressive_Fixation_Duration;
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

            private double m_Regressive_Saccade_Length;
            [Description("Regressive Saccade Length")]
            public double Regressive_Saccade_Length
            {
                get
                {
                    if (this.m_Regressive_Saccade_Length == -1)
                    {
                        this.m_Regressive_Saccade_Length = this.Regressive_Fixations.Sum(fix => fix.DistanceToPreviousFixation()) / this.Progressive_Fixations.Count;
                    }
                    return this.m_Regressive_Saccade_Length;
                }
            }

            private double m_Regressive_Saccade_Length_X;
            [Description("Regressive Saccade Length X")]
            public double Regressive_Saccade_Length_X
            {
                get
                {
                    if (this.m_Regressive_Saccade_Length_X == -1)
                    {
                        this.m_Regressive_Saccade_Length_X = this.Regressive_Fixations.Sum(fix => Math.Abs(fix.Fixation_Position_X - fix.Previous_Fixation.Fixation_Position_X)) / this.Progressive_Fixations.Count;
                    }
                    return this.m_Regressive_Saccade_Length_X;
                }
            }



            //[XLColumn(Ignore = true)]
            [EpplusIgnore]
            public List<Fixation> Fixations { get; set; }

            private List<Fixation> m_Progressive_Fixations;
            //[XLColumn(Ignore = true)]
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
            //[XLColumn(Ignore = true)]
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

            //private double m_StandardDeviation;
            //private double StandardDeviation
            //{
            //    get
            //    {
            //        if(m_StandardDeviation==-1)
            //        {
            //            m_StandardDeviation=
            //        }
            //        return m_StandardDeviation;
            //    }
            //}

            //private double m_Average;
            //private double StandardDeviation
            //{
            //    get
            //    {
            //        if (m_StandardDeviation == -1)
            //        {
            //            m_StandardDeviation =
            //        }
            //        return m_StandardDeviation;
            //    }
            //}

            //Average

            public ParticipantTrial(string Trial, string Stimulus, string Participant)//, List<Fixation> Fixations)//=null)
            {
                this.Trial = Trial;
                this.Stimulus = Stimulus;
                this.Participant = Participant;
                this.Fixations = FixationsService.fixationSetToFixationListDictionary[this.Participant + '\t' + this.Trial + '\t' + this.Stimulus];
                // Removes all the fixations with no AOI Group (after dealing with exceptions).
                this.Fixations.RemoveAll(fix => fix.AOI_Group_After_Change == -1);

                this.m_Total_Fixation_Number = -1;
                this.m_Mean_Fixation_Duration = -1;
                this.m_Progressive_Fixation_Duration = -1;
                this.m_Progressive_Saccade_Length = -1;
                this.m_Progressive_Saccade_Length_X = -1;
                this.m_Progressive_Fixations = null;
                this.m_Regressive_Fixation_Duration = -1;
                this.m_Regressive_Saccade_Length = -1;
                this.m_Regressive_Saccade_Length_X = -1;
                //this.m_StandardDeviation = -1;
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
