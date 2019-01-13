using ClosedXML.Attributes;
using ForBarIlanResearch.Enums;
using ForBarIlanResearch.ModelClasses;
using ForBarIlanResearch.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ForBarIlanResearch.ExcelsFilesMakers
{
    class SecondFileAfterProccessing
    {
        /// <summary>
        /// Makes the second excel file after proccessing.
        /// </summary>
        public static void makeExcelFile()
        {
            // Gets all the partisipants.
            List<string> participants = FixationsService.fixationSetToFixationListDictionary.Keys.ToList();
            string dictionatyKey;
            foreach (string participantKey in participants)
            {
                // Gets the current fixations list
                List<Fixation> fixations = FixationsService.fixationSetToFixationListDictionary[participantKey];
                int lastChangeIndex = 0, currentIndex = 0, last_AOIGroup = fixations[0].AOI_Group_After_Change, maxAOIGroupUntilNow = -1;
                Fixation prevFixationInAOI = null;

                foreach (Fixation fixation in fixations)
                {
                    //if (fixation.AOI_Name == 7)
                    //    Console.WriteLine(fixation.ToString());
                    if (fixation.AOI_Group_After_Change != last_AOIGroup)
                    {
                        // The dictionary key for the current AOI Group for the current Participant
                        dictionatyKey = participantKey + '\t' + last_AOIGroup;
                        List<Fixation> fixationRange = fixations.GetRange(lastChangeIndex, currentIndex - lastChangeIndex);

                        // If the current fixation is the first fixation of the current participant in the AOI so 
                        // create a new AOI Class for the excel table and add it to the AOIClass dictionary.
                        if (!AOIClass.instancesDictionary.ContainsKey(dictionatyKey))
                        {
                            AOIClass.instancesDictionary[dictionatyKey] =
                            new AOIClass(
                                prevFixationInAOI.Trial,
                                prevFixationInAOI.Stimulus,
                                prevFixationInAOI.Participant,
                                prevFixationInAOI.AOI_Group_After_Change,
                                //if the current fixation's AOI is not bigger then all the previous fixations so we skip it
                                prevFixationInAOI.AOI_Group_After_Change < maxAOIGroupUntilNow
                                );
                        }

                        // Adds the new fixation range (with the same AOI Group and the same participant and there is no 
                        // fixations in this range that have another AOI Group.
                        AOIClass.instancesDictionary[dictionatyKey].Fixations.Add(fixationRange);
                        last_AOIGroup = fixation.AOI_Group_After_Change;
                        lastChangeIndex = currentIndex;
                        if (maxAOIGroupUntilNow < last_AOIGroup)
                            maxAOIGroupUntilNow = last_AOIGroup;
                    }
                    else
                        fixation.Previous_Fixation = prevFixationInAOI;
                    prevFixationInAOI = fixation;
                    currentIndex++;
                }
            }
            // List<AOIClass> AOIClasses = AOIClass.instancesDictionary.Values.ToList();
            // AOIClasses.Sort((first, second) => first.dictionaryKey.CompareTo(second.dictionaryKey));
            // ExcelsService.CreateExcelFromStringTable(ConfigurationService.getValue(ConfigurationService.Second_Excel_File_Name), AOIClasses);
            ExcelsService.CreateExcelFromStringTable(ConfigurationService.getValue(ConfigurationService.Second_Excel_File_Name), AOIClass.instancesDictionary.Values.ToList());
        }


        private class AOIClass
        {
            [XLColumn(Ignore = true)]
            public static Dictionary<string, AOIClass> instancesDictionary = new Dictionary<string, AOIClass>();
            [XLColumn(Header = "Participant")]
            public string Participant { get; set; }
            [XLColumn(Header = "Trial")]
            public string Trial { get; set; }
            [XLColumn(Header = "Stimulus")]
            public string Stimulus { get; set; }
            [XLColumn(Header = "Text Name")]
            public string Text_Name
            {
                get { return FixationsService.textName; }
            }
            [XLColumn(Header = "AOI Group")]
            public int AOI_Group { get; set; }

            private double m_Total_Fixation_Duration;
            [XLColumn(Header = "Total Fixation Duration")]
            public double Total_Fixation_Duration
            {
                get
                {
                    if (this.m_Total_Fixation_Duration == -1)
                    {
                        this.m_Total_Fixation_Duration = this.Fixations.Sum(lst => lst.Sum(fix => fix.Event_Duration));
                    }
                    return this.m_Total_Fixation_Duration;
                }
            }

            private int m_Total_Fixation_Number;
            [XLColumn(Header = "Total Fixation Number")]
            public int Total_Fixation_Number
            {
                get
                {
                    if (this.m_Total_Fixation_Number == -1)
                    {
                        this.m_Total_Fixation_Number = this.Fixations.Sum(lst => lst.Count);
                    }
                    return this.m_Total_Fixation_Number;
                }
            }

            [XLColumn(Header = "First Fixation Duration")]
            public double First_Fixation_Duration
            {
                get
                {
                    try
                    {
                        return Fixations[0][0].Event_Duration;
                    }
                    catch
                    {
                        return 0;
                    }
                }
            }

            private double m_First_Pass_Duration;
            [XLColumn(Header = "First-Pass Duration")]
            public double First_Pass_Duration
            {
                get
                {
                    if (this.m_First_Pass_Duration == -1)
                    {
                        this.m_First_Pass_Duration = this.First_Pass_Fixations.Sum(fix => fix.Event_Duration);
                    }
                    return this.m_First_Pass_Duration;
                }
            }

            private int m_First_Pass_Number;
            [XLColumn(Header = "First-Pass Number")]
            public int First_Pass_Number
            {
                get
                {
                    if (this.m_First_Pass_Number == -1)
                    {
                        this.m_First_Pass_Number = this.First_Pass_Fixations.Count;
                    }
                    return this.m_First_Pass_Number;
                }
            }

            //private double m_First_Pass_Progressive_Duration;
            [XLColumn(Header = "First-Pass Progressive Duration")]
            public double First_Pass_Progressive_Duration
            {
                get
                {
                    return this.First_Pass_Progressive_Duration_Overall - this.First_Pass_Fixations[0].Event_Duration;
                }
            }

            [XLColumn(Header = "First-Pass Progressive Number")]
            public int First_Pass_Progressive_Number
            {
                get
                {
                    return this.First_Pass_Progressive_Number_Overall - 1;
                }
            }

            private double m_First_Pass_Progressive_Duration_Overall;
            [XLColumn(Header = "First-Pass Progressive Duration Overall")]
            public double First_Pass_Progressive_Duration_Overall
            {
                get
                {
                    if (this.m_First_Pass_Progressive_Duration_Overall == -1)
                    {
                        this.m_First_Pass_Progressive_Duration_Overall = this.Fixations_Progressive_First_Pass.Sum(fix => fix.Event_Duration);
                    }
                    return this.m_First_Pass_Progressive_Duration_Overall;
                }
            }

            [XLColumn(Header = "First-Pass Progressive Number Overall")]
            public int First_Pass_Progressive_Number_Overall
            {
                get
                {
                    return this.Fixations_Progressive_First_Pass.Count;
                }
            }

            [XLColumn(Header = "Total First-Pass Progressive Duration")]
            public double Total_First_Pass_Progressive_Duration
            {
                get
                {
                    return this.Total_First_Pass_Progressive_Duration_Overall - this.First_Pass_Fixations[0].Event_Duration;
                }
            }

            [XLColumn(Header = "Total First-Pass Progressive Number")]
            public int Total_First_Pass_Progressive_Number
            {
                get
                {
                    return this.Total_First_Pass_Progressive_Number_Overall - 1;
                }
            }

            private double m_Total_First_Pass_Progressive_Duration_Overall;
            [XLColumn(Header = "Total First-Pass Progressive Duration Overall")]
            public double Total_First_Pass_Progressive_Duration_Overall
            {
                get
                {
                    if (this.m_Total_First_Pass_Progressive_Duration_Overall == -1)
                    {
                        this.m_Total_First_Pass_Progressive_Duration_Overall = this.Total_Fixations_Progressive_First_Pass.Sum(fix => fix.Event_Duration);
                    }
                    return this.m_Total_First_Pass_Progressive_Duration_Overall;
                }
            }

            [XLColumn(Header = "Total First-Pass Progressive Number Overall")]
            public int Total_First_Pass_Progressive_Number_Overall
            {
                get
                {
                    return this.m_Total_Fixations_Progressive_First_Pass.Count;
                }
            }

            [XLColumn(Header = "Total First-Pass Regressive Duration")]
            public double Total_First_Pass_Regressive_Duration
            {
                get
                {
                    return this.First_Pass_Duration - this.Total_First_Pass_Progressive_Duration_Overall;
                }
            }
            [XLColumn(Header = "Total First-Pass Regressive Number")]
            public int Total_First_Pass_Regressive_Number
            {
                get
                {
                    return this.First_Pass_Number - this.Total_First_Pass_Progressive_Number_Overall;
                }
            }

            [XLColumn(Header = "Regression Number")]
            public int Regression_Number
            {
                get
                {
                    return this.Fixations.Count - 1;
                }
            }
            [XLColumn(Header = "Regression Duration")]
            public double Regression_Duration
            {
                get
                {
                    return this.Total_Fixation_Duration - this.First_Pass_Duration;
                }
            }

            private double m_First_Regression_Duration;
            [XLColumn(Header = "First Regression Duration")]
            public double First_Regression_Duration
            {
                get
                {
                    if (this.m_First_Regression_Duration == -1)
                    {
                        if (this.Fixations.Count > 1)
                            this.m_First_Regression_Duration = this.Fixations[1].Sum(fix => fix.Event_Duration);
                        else
                            this.m_First_Regression_Duration = 0;
                    }
                    return this.m_First_Regression_Duration;
                }
            }

            [XLColumn(Header = "Skip")]
            public bool Skip { get; set; }
            [XLColumn(Header = "Pupil Diameter [mm]")]
            public double Pupil_Diameter
            {
                get
                {
                    return this.Total_Pupil_Diameter / this.Total_Fixation_Number;
                }
            }
            [XLColumn(Header = "AOI Size X [mm]")]
            public double Mean_AOI_Size
            {
                get
                {
                    return this.Total_AOI_Size / this.Total_Fixation_Number;
                }
            }
            [XLColumn(Header = "AOI Converage [%]")]
            public double Mean_AOI_Converage
            {
                get
                {
                    return this.Total_AOI_Converage / this.Total_Fixation_Number;
                }
            }

            private double m_Total_AOI_Size;
            [XLColumn(Ignore = true)]
            public double Total_AOI_Size
            {
                get
                {
                    if (this.m_Total_AOI_Size == -1)
                    {
                        //this.m_Total_AOI_Size = this.Fixations.Sum(lst => lst.Sum(fix => fix.AOI_Size));
                    }
                    return this.m_Total_AOI_Size;
                }
            }

            private double m_Total_AOI_Converage;
            [XLColumn(Ignore = true)]
            public double Total_AOI_Converage
            {
                get
                {
                    if (this.m_Total_AOI_Converage == -1)
                    {
                        this.m_Total_AOI_Converage = this.Fixations.Sum(lst => lst.Sum(fix => fix.AOI_Coverage_In_Percents));
                    }
                    return this.m_Total_AOI_Converage;
                }
            }

            private double m_Total_Pupil_Diameter;
            [XLColumn(Ignore = true)]
            public double Total_Pupil_Diameter
            {
                get
                {
                    if (this.m_Total_Pupil_Diameter == -1)
                    {
                        this.m_Total_Pupil_Diameter = this.Fixations.Sum(lst => lst.Sum(fix => fix.Fixation_Average_Pupil_Diameter));
                    }
                    return this.m_Total_Pupil_Diameter;
                }
            }

            [XLColumn(Ignore = true)]
            public List<List<Fixation>> Fixations { get; set; }

            private List<Fixation> m_First_Pass_Fixations;
            [XLColumn(Ignore = true)]
            public List<Fixation> First_Pass_Fixations
            {
                get
                {
                    try
                    {
                        if (this.m_First_Pass_Fixations == null)
                        {
                            this.m_First_Pass_Fixations = this.Fixations.First();
                            DealingWithExceptionsEnum dealingWithInsideExceptions = (DealingWithExceptionsEnum)int.Parse(ConfigurationService.getValue(ConfigurationService.Dealing_With_Exceptions_Inside_The_Limit));
                            DealingWithExceptionsOutBoundsEnum dealingWithOutsideExceptions = (DealingWithExceptionsOutBoundsEnum)int.Parse(ConfigurationService.getValue(ConfigurationService.Dealing_With_Exceptions_Outside_The_Limit));
                            if (dealingWithInsideExceptions== DealingWithExceptionsEnum.Skip_In_First_Pass)
                            {
                                this.m_First_Pass_Fixations.RemoveAll(fix => fix.IsException && fix.IsInExceptionBounds);
                            }
                            if(dealingWithOutsideExceptions==DealingWithExceptionsOutBoundsEnum.Skip_In_First_Pass)
                            {
                                this.m_First_Pass_Fixations.RemoveAll(fix => fix.IsException && !fix.IsInExceptionBounds);
                            }
                        }
                        return this.m_First_Pass_Fixations;
                    }
                    catch
                    {
                        return null;
                    }
                }
            }

            private List<Fixation> m_Fixations_Progressive_First_Pass;
            private List<Fixation> Fixations_Progressive_First_Pass
            {
                get
                {
                    if (this.m_Fixations_Progressive_First_Pass == null)
                    {
                        Fixation firstFixation = this.First_Pass_Fixations[0];

                        //int firstOneRightToFirst = this.First_Pass_Fixations.FindIndex(fix => fix.isBeforeThan(firstFixation));
                        int firstOneRightToFirst = this.First_Pass_Fixations.FindIndex(fix => fix.Previous_Fixation!=null && fix.isBeforeThan(fix.Previous_Fixation));
                        if (firstOneRightToFirst == -1)
                            firstOneRightToFirst = this.First_Pass_Fixations.Count;
                        this.m_Fixations_Progressive_First_Pass = this.First_Pass_Fixations.GetRange(0, firstOneRightToFirst);
                    }
                    return this.m_Fixations_Progressive_First_Pass;
                }
            }

            private List<Fixation> m_Total_Fixations_Progressive_First_Pass;
            private List<Fixation> Total_Fixations_Progressive_First_Pass
            {
                get
                {
                    if (this.m_Total_Fixations_Progressive_First_Pass == null)
                    {
                        this.m_Total_Fixations_Progressive_First_Pass = new List<Fixation>();
                        Fixation[] first_Pass_Fixations = this.First_Pass_Fixations.ToArray();
                        this.m_Total_Fixations_Progressive_First_Pass.Add(first_Pass_Fixations[0]);
                        for (int i = 1; i < first_Pass_Fixations.Length; ++i)
                        {
                            if (first_Pass_Fixations[i - 1].isBeforeThan(first_Pass_Fixations[i]))
                                this.m_Total_Fixations_Progressive_First_Pass.Add(first_Pass_Fixations[i]);
                        }
                    }
                    return this.m_Fixations_Progressive_First_Pass;
                }
            }

            // [XLColumn(Ignore = true)]
            // public readonly string dictionaryKey;

            public AOIClass(string Trial, string Stimulus, string Participant, int AOI_Group, bool Skip)
            {
                this.Trial = Trial;
                this.Stimulus = Stimulus;
                this.Participant = Participant;
                this.AOI_Group = AOI_Group;
                this.Skip = Skip;
                this.Fixations = new List<List<Fixation>>();


                this.m_Total_Fixation_Duration = -1;
                this.m_Total_Fixation_Number = -1;
                this.m_First_Pass_Duration = -1;
                this.m_First_Pass_Number = -1;
                this.m_Fixations_Progressive_First_Pass = null;
                this.m_First_Pass_Fixations = null;
                this.m_First_Pass_Progressive_Duration_Overall = -1;
                this.m_Total_First_Pass_Progressive_Duration_Overall = -1;
                this.m_First_Regression_Duration = -1;
                this.m_Total_Pupil_Diameter = -1;
                this.m_Total_AOI_Size = -1;
                this.m_Total_AOI_Converage = -1;

            }
            
            public string GetDictionaryKey()
            {
                return this.Participant + '\t' + this.Trial + '\t' + this.Stimulus + '\t' + this.AOI_Group;
            }

        }
    }
}
