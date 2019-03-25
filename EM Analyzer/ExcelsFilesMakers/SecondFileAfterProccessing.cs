using ClosedXML.Attributes;
using EM_Analyzer.Enums;
using EM_Analyzer.ModelClasses;
using EM_Analyzer.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EM_Analyzer.ExcelsFilesMakers
{
    class SecondFileAfterProccessing
    {

        private delegate double NumericExpression(AOIClass value);
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
                fixations.Add(new Fixation() { AOI_Group_After_Change = -2 });
                int lastChangeIndex = 0, currentIndex = 0, last_AOIGroup = fixations[0].AOI_Group_After_Change, maxAOIGroupUntilNow = -1;
                Fixation prevFixationInAOI = null;

                foreach (Fixation fixation in fixations)
                {
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
                fixations.RemoveAt(fixations.Count - 1);
            }
            // List<AOIClass> AOIClasses = AOIClass.instancesDictionary.Values.ToList();
            // AOIClasses.Sort((first, second) => first.dictionaryKey.CompareTo(second.dictionaryKey));
            // ExcelsService.CreateExcelFromStringTable(ConfigurationService.getValue(ConfigurationService.Second_Excel_File_Name), AOIClasses);
            ExcelsService.CreateExcelFromStringTable(ConfigurationService.SecondExcelFileName, AOIClass.instancesDictionary.Values.ToList());
            AOIClass.considerCoverage = true;
            List<NumericExpression> filteringsExpressions = new List<NumericExpression>();
            filteringsExpressions.Add(aoi => aoi.Total_Fixation_Duration);
            filteringsExpressions.Add(aoi => aoi.Total_Fixation_Number);
            filteringsExpressions.Add(aoi => aoi.First_Fixation_Duration);
            filteringsExpressions.Add(aoi => aoi.First_Pass_Duration);
            filteringsExpressions.Add(aoi => aoi.First_Pass_Number);
            filteringsExpressions.Add(aoi => aoi.First_Pass_Progressive_Duration);
            filteringsExpressions.Add(aoi => aoi.First_Pass_Progressive_Number);
            filteringsExpressions.Add(aoi => aoi.First_Pass_Progressive_Duration_Overall);
            filteringsExpressions.Add(aoi => aoi.First_Pass_Progressive_Number_Overall);
            filteringsExpressions.Add(aoi => aoi.Total_First_Pass_Progressive_Duration);
            filteringsExpressions.Add(aoi => aoi.Total_First_Pass_Progressive_Number);
            filteringsExpressions.Add(aoi => aoi.Total_First_Pass_Progressive_Duration_Overall);
            filteringsExpressions.Add(aoi => aoi.Total_First_Pass_Progressive_Number_Overall);
            filteringsExpressions.Add(aoi => aoi.Total_First_Pass_Regressive_Duration);
            filteringsExpressions.Add(aoi => aoi.Total_First_Pass_Regressive_Number);
            filteringsExpressions.Add(aoi => aoi.Regression_Number);
            filteringsExpressions.Add(aoi => aoi.Regression_Duration);
            filteringsExpressions.Add(aoi => aoi.First_Regression_Duration);
            filteringsExpressions.Add(aoi => aoi.Pupil_Diameter);
            ExcelsService.CreateExcelFromStringTable(ConfigurationService.ConsideredSecondExcelFileName, AOIClass.instancesDictionary.Values.ToList());
        }


        private class AOIClass
        {
            private static readonly uint minimumNumberOfFixationsInARegression= uint.Parse(ConfigurationService.MinimumNumberOfFixationsInARegression);
            [XLColumn(Ignore = true)]
            public static Dictionary<string, AOIClass> instancesDictionary = new Dictionary<string, AOIClass>();

            public static bool considerCoverage = false;

            [XLColumn(Header = "Participant")]
            public string Participant { get; set; }
            [XLColumn(Header = "Trial")]
            public string Trial { get; set; }
            [XLColumn(Header = "Stimulus")]
            public string Stimulus { get; set; }
            [XLColumn(Header = "Text Name")]
            public string Text_Name
            {
                get { return FixationsService.excelFileName; }
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
                    if(considerCoverage)
                        return this.m_Total_Fixation_Duration/this.Mean_AOI_Coverage;
                    return this.m_Total_Fixation_Duration;
                }
            }

            private double m_Total_Fixation_Number;
            [XLColumn(Header = "Total Fixation Number")]
            public double Total_Fixation_Number
            {
                get
                {
                    if (this.m_Total_Fixation_Number == -1)
                    {
                        this.m_Total_Fixation_Number = this.Fixations.Sum(lst => lst.Count);
                    }
                    if (considerCoverage)
                        return this.m_Total_Fixation_Number / this.Mean_AOI_Coverage;
                    return this.m_Total_Fixation_Number;
                }
            }

            private double m_First_Fixation_Duration=-1;
            [XLColumn(Header = "First Fixation Duration")]
            public double First_Fixation_Duration
            {
                get
                {
                    if (m_First_Fixation_Duration == -1)
                    {
                        try
                        {
                            m_First_Fixation_Duration = Fixations[0][0].Event_Duration;
                        }
                        catch
                        {
                            return 0;
                        }
                    }

                    if (considerCoverage)
                        return this.m_First_Fixation_Duration / this.Mean_AOI_Coverage;
                    return m_First_Fixation_Duration;
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

                    if (considerCoverage)
                        return this.m_First_Pass_Duration / this.Mean_AOI_Coverage;
                    return this.m_First_Pass_Duration;
                }
            }

            private double m_First_Pass_Number;
            [XLColumn(Header = "First-Pass Number")]
            public double First_Pass_Number
            {
                get
                {
                    if (this.m_First_Pass_Number == -1)
                    {
                        this.m_First_Pass_Number = this.First_Pass_Fixations.Count;
                    }

                    if (considerCoverage)
                        return this.m_First_Pass_Number / this.Mean_AOI_Coverage;
                    return this.m_First_Pass_Number;
                }
            }

            private double m_First_Pass_Progressive_Duration=-1;
            [XLColumn(Header = "First-Pass Progressive Duration")]
            public double First_Pass_Progressive_Duration
            {
                get
                {
                    if(m_First_Pass_Progressive_Duration==-1)
                        m_First_Pass_Progressive_Duration = this.First_Pass_Progressive_Duration_Overall - this.First_Pass_Fixations[0].Event_Duration;

                    if (considerCoverage)
                        return this.m_First_Pass_Progressive_Duration / this.Mean_AOI_Coverage;
                    return m_First_Pass_Progressive_Duration;
                }
            }

            private double m_First_Pass_Progressive_Number=-1;
            [XLColumn(Header = "First-Pass Progressive Number")]
            public double First_Pass_Progressive_Number
            {
                get
                {
                    if(m_First_Pass_Progressive_Number==-1)
                        m_First_Pass_Progressive_Number = this.First_Pass_Progressive_Number_Overall - 1;

                    if (considerCoverage)
                        return this.m_First_Pass_Progressive_Number / this.Mean_AOI_Coverage;
                    return m_First_Pass_Progressive_Number;
                }
            }

            private double m_First_Pass_Progressive_Duration_Overall=-1;
            [XLColumn(Header = "First-Pass Progressive Duration Overall")]
            public double First_Pass_Progressive_Duration_Overall
            {
                get
                {
                    if (this.m_First_Pass_Progressive_Duration_Overall == -1)
                    {
                        this.m_First_Pass_Progressive_Duration_Overall = this.Fixations_Progressive_First_Pass.Sum(fix => fix.Event_Duration);
                    }

                    if (considerCoverage)
                        return this.m_First_Pass_Progressive_Duration_Overall / this.Mean_AOI_Coverage;
                    return this.m_First_Pass_Progressive_Duration_Overall;
                }
            }

            private double m_First_Pass_Progressive_Number_Overall =-1;
            [XLColumn(Header = "First-Pass Progressive Number Overall")]
            public double First_Pass_Progressive_Number_Overall
            {
                get
                {
                    if(m_First_Pass_Progressive_Number_Overall==-1)
                        m_First_Pass_Progressive_Number_Overall= this.Fixations_Progressive_First_Pass.Count;

                    if (considerCoverage)
                        return this.m_First_Pass_Progressive_Number_Overall / this.Mean_AOI_Coverage;
                    return m_First_Pass_Progressive_Number_Overall;
                }
            }

            private double m_Total_First_Pass_Progressive_Duration = -1;
            [XLColumn(Header = "Total First-Pass Progressive Duration")]
            public double Total_First_Pass_Progressive_Duration
            {
                get
                {
                    if(m_Total_First_Pass_Progressive_Duration==-1)
                        m_Total_First_Pass_Progressive_Duration= this.Total_First_Pass_Progressive_Duration_Overall - this.First_Pass_Fixations[0].Event_Duration;

                    if (considerCoverage)
                        return this.m_Total_First_Pass_Progressive_Duration / this.Mean_AOI_Coverage;
                    return m_Total_First_Pass_Progressive_Duration;
                }
            }

            private double m_Total_First_Pass_Progressive_Number = -1;
            [XLColumn(Header = "Total First-Pass Progressive Number")]
            public double Total_First_Pass_Progressive_Number
            {
                get
                {
                    if(m_Total_First_Pass_Progressive_Number==-1)
                        m_Total_First_Pass_Progressive_Number= this.Total_First_Pass_Progressive_Number_Overall - 1;

                    if (considerCoverage)
                        return this.m_Total_First_Pass_Progressive_Number / this.Mean_AOI_Coverage;
                    return m_Total_First_Pass_Progressive_Number;
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

                    if (considerCoverage)
                        return this.m_Total_First_Pass_Progressive_Duration_Overall / this.Mean_AOI_Coverage;
                    return this.m_Total_First_Pass_Progressive_Duration_Overall;
                }
            }

            private double m_Total_First_Pass_Progressive_Number_Overall = -1;
            [XLColumn(Header = "Total First-Pass Progressive Number Overall")]
            public double Total_First_Pass_Progressive_Number_Overall
            {
                get
                {
                    if(m_Total_First_Pass_Progressive_Number_Overall==-1)
                        m_Total_First_Pass_Progressive_Number_Overall= this.m_Total_Fixations_Progressive_First_Pass.Count;

                    if (considerCoverage)
                        return this.m_Total_First_Pass_Progressive_Number_Overall / this.Mean_AOI_Coverage;
                    return m_Total_First_Pass_Progressive_Number_Overall;
                }
            }

            private double m_Total_First_Pass_Regressive_Duration=-1;
            [XLColumn(Header = "Total First-Pass Regressive Duration")]
            public double Total_First_Pass_Regressive_Duration
            {
                get
                {
                    if(m_Total_First_Pass_Regressive_Duration==-1)
                        m_Total_First_Pass_Regressive_Duration= this.First_Pass_Duration - this.Total_First_Pass_Progressive_Duration_Overall;

                    if (considerCoverage)
                        return this.m_Total_First_Pass_Regressive_Duration / this.Mean_AOI_Coverage;
                    return m_Total_First_Pass_Regressive_Duration;
                }
            }

            private double m_Total_First_Pass_Regressive_Number = -1;
            [XLColumn(Header = "Total First-Pass Regressive Number")]
            public double Total_First_Pass_Regressive_Number
            {
                get
                {
                    if(m_Total_First_Pass_Regressive_Number==-1)
                        m_Total_First_Pass_Regressive_Number= this.First_Pass_Number - this.Total_First_Pass_Progressive_Number_Overall;

                    if (considerCoverage)
                        return this.m_Total_First_Pass_Regressive_Number / this.Mean_AOI_Coverage;
                    return m_Total_First_Pass_Regressive_Number;
                }
            }

            private double m_Regression_Number = -1;
            [XLColumn(Header = "Regression Number")]
            public double Regression_Number
            {
                get
                {
                    if(m_Regression_Number==-1)
                        m_Regression_Number=this.Regressions.Count();

                    if (considerCoverage)
                        return this.m_Regression_Number / this.Mean_AOI_Coverage;
                    return m_Regression_Number;
                }
            }

            private double m_Regression_Duration = -1;
            [XLColumn(Header = "Regression Duration")]
            public double Regression_Duration
            {
                get
                {
                    //return this.Total_Fixation_Duration - this.First_Pass_Duration;
                    if(m_Regression_Duration==-1)
                        m_Regression_Duration= this.Regressions.Sum(lst => lst.Sum(fix => fix.Event_Duration));

                    if (considerCoverage)
                        return this.m_Regression_Duration / this.Mean_AOI_Coverage;
                    return m_Regression_Duration;
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
                        if (this.Regressions.Any())
                            this.m_First_Regression_Duration = this.Regressions.First().Sum(fix => fix.Event_Duration);
                        else
                            this.m_First_Regression_Duration = 0;
                    }

                    if (considerCoverage)
                        return this.m_First_Regression_Duration / this.Mean_AOI_Coverage;
                    return this.m_First_Regression_Duration;
                }
            }

            [XLColumn(Header = "Skip")]
            public bool Skip { get; set; }

            private double m_Pupil_Diameter = -1;
            [XLColumn(Header = "Pupil Diameter [mm]")]
            public double Pupil_Diameter
            {
                get
                {
                    if(m_Pupil_Diameter==-1)
                        m_Pupil_Diameter= this.Total_Pupil_Diameter / this.Total_Fixation_Number;

                    if (considerCoverage)
                        return this.m_Pupil_Diameter / this.Mean_AOI_Coverage;
                    return m_Pupil_Diameter;
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

            private double m_Mean_AOI_Coverage = -1;
            [XLColumn(Header = "AOI Coverage [%]")]
            public double Mean_AOI_Coverage
            {
                get
                {
                    if(m_Mean_AOI_Coverage==-1)
                        m_Mean_AOI_Coverage= this.Total_AOI_Coverage / this.Total_Fixation_Number;
                    return m_Mean_AOI_Coverage;
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
                        this.m_Total_AOI_Size = this.Fixations.Sum(lst => lst.Sum(fix => fix.AOI_Size));
                    }
                    return this.m_Total_AOI_Size;
                }
            }

            private double m_Total_AOI_Coverage;
            [XLColumn(Ignore = true)]
            public double Total_AOI_Coverage
            {
                get
                {
                    if (this.m_Total_AOI_Coverage == -1)
                    {
                        this.m_Total_AOI_Coverage = this.Fixations.Sum(lst => lst.Sum(fix => fix.AOI_Coverage_In_Percents));
                    }
                    return this.m_Total_AOI_Coverage;
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
                            DealingWithExceptionsEnum dealingWithInsideExceptions = (DealingWithExceptionsEnum)int.Parse(ConfigurationService.DealingWithExceptionsInsideTheLimit);
                            DealingWithExceptionsOutBoundsEnum dealingWithOutsideExceptions = (DealingWithExceptionsOutBoundsEnum)int.Parse(ConfigurationService.DealingWithExceptionsOutsideTheLimit);
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
                    return this.m_Total_Fixations_Progressive_First_Pass;
                }
            }

            private IEnumerable<List<Fixation>> m_Regressions;
            private IEnumerable<List<Fixation>> Regressions
            {
                get
                {
                    if (this.m_Regressions == null)
                    {
                        this.m_Regressions= this.Fixations.GetRange(1, this.Fixations.Count - 1);
                        this.m_Regressions = this.m_Regressions.Where(lst => lst.Count >= minimumNumberOfFixationsInARegression);
                    }
                    return this.m_Regressions;
                }
            }

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
                this.m_Regressions = null;
                this.m_First_Pass_Progressive_Duration_Overall = -1;
                this.m_Total_First_Pass_Progressive_Duration_Overall = -1;
                this.m_First_Regression_Duration = -1;
                this.m_Total_Pupil_Diameter = -1;
                this.m_Total_AOI_Size = -1;
                this.m_Total_AOI_Coverage = -1;

            }
            
            public string GetDictionaryKey()
            {
                return this.Participant + '\t' + this.Trial + '\t' + this.Stimulus + '\t' + this.AOI_Group;
            }

        }
    }
}
