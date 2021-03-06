using EM_Analyzer.Enums;
using EM_Analyzer.ModelClasses;
using EM_Analyzer.Services;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using EM_Analyzer.Interfaces;
using EM_Analyzer.ModelClasses.AOIClasses;
using System.Dynamic;

namespace EM_Analyzer.ExcelsFilesMakers
{
    public class SecondFileAfterProccessing
    {
        //public delegate double NumericExpression(AOIClass value);
        /// <summary>
        /// Makes the second excel file after proccessing
        /// </summary>
        public static AOITypes currentType;
        public static void MakeExcelFile()
        {
            IEnumerable<IGrouping<string, KeyValuePair<string, List<Fixation>>>> fixationsGroupingByParticipant = FixationsService.fixationSetToFixationListDictionary.GroupBy(x => x.Value[0].Participant);
            string dictionatyKey;
            //int minimumEventDurationInForSkipInms = int.Parse(ConfigurationService.MinimumEventDurationInForSkipInms);
            int minimumNumberOfFixationsForSkip = int.Parse(ConfigurationService.MinimumNumberOfFixationsForSkip);
            List<string> dictionaryKeysForSorting = new List<string>();
            foreach (IGrouping<string, KeyValuePair<string, List<Fixation>>> participantFixations in fixationsGroupingByParticipant)
            {
                string participantKey = participantFixations.Key;
                // Gets the current fixations list
                List<Fixation> fixations = participantFixations.SelectMany(x => x.Value).ToList();

                // remove the fixations in figure object?
                List<Fixation> fixationsForFirstPass = fixations.ToList();
                // Make filter per fixation
                fixationsForFirstPass.RemoveAll(fix => fix.ShouldBeSkippedInFirstPass());
                List<CountedAOIFixations> countedAOIFixationsForFirstPass;
                if (currentType == AOITypes.Phrases)
                    countedAOIFixationsForFirstPass = FixationsService.ConvertFixationListToCoutedListByPhrase(fixationsForFirstPass);
                else
                    countedAOIFixationsForFirstPass = FixationsService.ConvertFixationListToCoutedListByWords(fixationsForFirstPass);

                #region First_Pass
                // For The First Pass Fixations
                foreach (CountedAOIFixations countedAOIFixations in countedAOIFixationsForFirstPass)
                {
                    // First try with first pass
                    if (!FixationsService.IsLeagalFirstPassFixations(countedAOIFixations) && currentType == AOITypes.Phrases)
                        continue;
                    // End

                    dictionatyKey = participantKey + '\t' + countedAOIFixations.AOI_Group;
                    if (!AOIClass.instancesDictionary.ContainsKey(dictionatyKey))
                    {
                        Fixation fixationInAOI = countedAOIFixations.Fixations.First();
                        AOIClass.instancesDictionary[dictionatyKey] =
                            new AOIClass(
                                fixationInAOI.Stimulus,
                                fixationInAOI.Participant,
                                countedAOIFixations.AOI_Group,
                                false,
                                currentType
                                //if the current fixation's AOI is not bigger then all the previous fixations so we skip it
                                //prevFixationInAOI.AOI_Group_After_Change < maxAOIGroupUntilNow
                                )
                            {
                                // maybe here, check the size if the counted aoi fixation, minimum number of fixations for first pass
                                // and minimum duration of fixation for first pass
                                First_Pass_Fixations = countedAOIFixations.Fixations
                            };
                
                    }
                }
                #endregion First_Pass
                // For All The Rest
                fixations.Add(new Fixation() { AOI_Group_After_Change = -2 });
                int lastChangeIndex = 0, currentIndex = 0, last_AOI = GetAOIByType(fixations[0]), maxLeagalAOIGroupUntilNow = -1;
                Fixation prevFixationInAOI = null;
                if (currentType == AOITypes.Words)
                {
                    // clear previous fixations from the calculations by phrases
                    FixationsService.ClearPreviousFixation(fixations);
                }
                foreach (Fixation fixation in fixations)
                {
                    
                    if (GetAOIByType(fixation) != last_AOI)
                    {
                        // The dictionary key for the current AOI Group for the current Participant
                        dictionatyKey = participantKey + '\t' + last_AOI;
                        

                        List<Fixation> fixationRange = fixations.GetRange(lastChangeIndex, currentIndex - lastChangeIndex);

                        // If the current fixation is the first fixation of the current participant in the AOI so 
                        // create a new AOI Class for the excel table and add it to the AOIClass dictionary.
                        if (!AOIClass.instancesDictionary.ContainsKey(dictionatyKey))
                        {
                            AOIClass.instancesDictionary[dictionatyKey] =
                            new AOIClass(
                                prevFixationInAOI.Stimulus,
                                prevFixationInAOI.Participant,
                                GetAOIByType(prevFixationInAOI),
                                //if the current fixation's AOI is not greater then all the previous fixations so we skip it
                                GetAOIByType(prevFixationInAOI) < maxLeagalAOIGroupUntilNow,
                                currentType
                                );
                        }
                        else
                        {
                            if (!dictionaryKeysForSorting.Contains(dictionatyKey))
                                AOIClass.instancesDictionary[dictionatyKey].Skip = GetAOIByType(prevFixationInAOI) < maxLeagalAOIGroupUntilNow;
                        }
                        if (!dictionaryKeysForSorting.Contains(dictionatyKey))
                            dictionaryKeysForSorting.Add(dictionatyKey);

                        // Adds the new fixation range (with the same AOI Group and the same participant and there is no 
                        // fixations in this range that have another AOI Group.
                        AOIClass.instancesDictionary[dictionatyKey].Fixations.Add(fixationRange);

                        if (maxLeagalAOIGroupUntilNow < last_AOI && 
                            (FixationsService.IsLeagalFixationsForSkip(fixationRange) || currentType == AOITypes.Words))
                            maxLeagalAOIGroupUntilNow = last_AOI;

                        last_AOI = GetAOIByType(fixation);
                        lastChangeIndex = currentIndex;

                    }
                    else
                        fixation.Previous_Fixation = prevFixationInAOI;
                    prevFixationInAOI = fixation;
                    currentIndex++;
                }
                fixations.RemoveAt(fixations.Count - 1);
            }
            List<AOIClass> aoiClasses = dictionaryKeysForSorting.Select(key=> AOIClass.instancesDictionary[key]).ToList();
            
            // not to show the figure in words output
            if (currentType == AOITypes.Words)
                aoiClasses.RemoveAll(aoi => aoi.AOI_Group == 0);
            
            ExcelsService.CreateExcelFromStringTable(
                ConfigurationService.SecondExcelFileName + "_" + Constans.GetEndOfFileNameByType(currentType), aoiClasses,
              SecondFileConsideringCoverage.EditExcel); 
        }
        
        public static int GetAOIByType(Fixation fixation)
        {
            if (currentType == AOITypes.Phrases)
                return fixation.AOI_Group_After_Change;
            else
                return fixation.Word_Index;
        }

        public class AOIClass : IAOIClassForConsideringCoverage
        {
            private static readonly uint minimumNumberOfFixationsInARegression = uint.Parse(ConfigurationService.MinimumNumberOfFixationsForRegression);
            public static Dictionary<string, AOIClass> instancesDictionary = new Dictionary<string, AOIClass>();
            public static int maxConditions;

            [Description("Participant")]
            public string Participant { get; set; }
            [Description("Stimulus")]
            public string Stimulus { get; set; }
            [Description("Text Name")]
            public string Text_Name
            {
                get { return FixationsService.phrasesExcelFileName; }
            }
            [Description("AOI Group")]
            public int AOI_Group { get; set; }
            [Description("AOI Target")]
            public string AOI_Target { get; set; }

            private double m_Total_Fixation_Duration;
            [Description("Total Fixation Duration")]
            public double Total_Fixation_Duration
            {
                get
                {
                    if (this.m_Total_Fixation_Duration == -1)
                        this.m_Total_Fixation_Duration = this.Fixations.Sum(lst => lst.Sum(fix => fix.Event_Duration));
                    return this.m_Total_Fixation_Duration;
                }
            }

            private double m_Total_Fixation_Number;
            [Description("Total Fixation Number")]
            public double Total_Fixation_Number
            {
                get
                {
                    if (this.m_Total_Fixation_Number == -1)
                        this.m_Total_Fixation_Number = this.Fixations.Sum(lst => lst.Count);
                    return this.m_Total_Fixation_Number;
                }
            }

            private double m_First_Fixation_Duration = -1;
            [Description("First Fixation Duration")]
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
                    return m_First_Fixation_Duration;
                }
            }

            private double m_First_Pass_Duration;
            [Description("First-Pass Duration")]
            public double First_Pass_Duration
            {
                get
                {
                    if (this.m_First_Pass_Duration == -1)
                        this.m_First_Pass_Duration = this.First_Pass_Fixations.Sum(fix => fix.Event_Duration);
                    return this.m_First_Pass_Duration;
                }
            }

            private double m_First_Pass_Number;
            [Description("First-Pass Number")]
            public double First_Pass_Number
            {
                get
                {
                    if (this.m_First_Pass_Number == -1)
                        this.m_First_Pass_Number = this.First_Pass_Fixations.Count;
                    return this.m_First_Pass_Number;
                }
            }

            private double m_First_Pass_Progressive_Duration = -1;
            [Description("First-Pass Progressive Duration")]
            public double First_Pass_Progressive_Duration
            {
                get
                {
                    if (m_First_Pass_Progressive_Duration == -1)
                    {
                        if (this.First_Pass_Fixations.Any())
                            m_First_Pass_Progressive_Duration = this.First_Pass_Progressive_Duration_Overall - this.First_Pass_Fixations[0].Event_Duration;
                        else
                        {
                            m_First_Pass_Progressive_Duration = 0;
                        }
                    }
                    return Math.Max(0, m_First_Pass_Progressive_Duration);
                }
            }

            private double m_First_Pass_Progressive_Number = -1;
            [Description("First-Pass Progressive Number")]
            public double First_Pass_Progressive_Number
            {
                get
                {
                    if (m_First_Pass_Progressive_Number == -1)
                        m_First_Pass_Progressive_Number = this.First_Pass_Progressive_Number_Overall - 1;
                    return Math.Max(0,m_First_Pass_Progressive_Number);
                }
            }

            private double m_First_Pass_Progressive_Duration_Overall = -1;
            [Description("First-Pass Progressive Duration Overall")]
            public double First_Pass_Progressive_Duration_Overall
            {
                get
                {
                    if (this.m_First_Pass_Progressive_Duration_Overall == -1)
                        this.m_First_Pass_Progressive_Duration_Overall = this.Fixations_Progressive_First_Pass.Sum(fix => fix.Event_Duration);
                    return Math.Max(0, m_First_Pass_Progressive_Duration_Overall);
                }
            }

            private double m_First_Pass_Progressive_Number_Overall = -1;
            [Description("First-Pass Progressive Number Overall")]
            public double First_Pass_Progressive_Number_Overall
            {
                get
                {
                    if (m_First_Pass_Progressive_Number_Overall == -1)
                        m_First_Pass_Progressive_Number_Overall = this.Fixations_Progressive_First_Pass.Count;
                    return m_First_Pass_Progressive_Number_Overall;
                }
            }

            private double m_Total_First_Pass_Progressive_Duration = -1;
            [Description("Total First-Pass Progressive Duration")]
            public double Total_First_Pass_Progressive_Duration
            {
                get
                {
                    if (m_Total_First_Pass_Progressive_Duration == -1)
                        if (this.First_Pass_Fixations.Any())
                            m_Total_First_Pass_Progressive_Duration = this.Total_First_Pass_Progressive_Duration_Overall - this.First_Pass_Fixations[0].Event_Duration;
                        else
                            m_Total_First_Pass_Progressive_Duration = 0;
                    return m_Total_First_Pass_Progressive_Duration;
                }
            }

            private double m_Total_First_Pass_Progressive_Number = -1;
            [Description("Total First-Pass Progressive Number")]
            public double Total_First_Pass_Progressive_Number
            {
                get
                {
                    if (m_Total_First_Pass_Progressive_Number == -1)
                        m_Total_First_Pass_Progressive_Number = this.Total_First_Pass_Progressive_Number_Overall - 1;
                    return Math.Max(0, m_Total_First_Pass_Progressive_Number);
                }
            }

            private double m_Total_First_Pass_Progressive_Duration_Overall;
            [Description("Total First-Pass Progressive Duration Overall")]
            public double Total_First_Pass_Progressive_Duration_Overall
            {
                get
                {
                    if (this.m_Total_First_Pass_Progressive_Duration_Overall == -1)
                        this.m_Total_First_Pass_Progressive_Duration_Overall = this.Total_Fixations_Progressive_First_Pass.Sum(fix => fix.Event_Duration);
                    return Math.Max(0, m_Total_First_Pass_Progressive_Duration_Overall);
                }
            }

            private double m_Total_First_Pass_Progressive_Number_Overall = -1;
            [Description("Total First-Pass Progressive Number Overall")]
            public double Total_First_Pass_Progressive_Number_Overall
            {
                get
                {
                    if (m_Total_First_Pass_Progressive_Number_Overall == -1)
                        m_Total_First_Pass_Progressive_Number_Overall = this.Total_Fixations_Progressive_First_Pass.Count;
                    return Math.Max(0, m_Total_First_Pass_Progressive_Number_Overall);
                }
            }

            private double m_Total_First_Pass_Regressive_Duration = -1;
            [Description("Total First-Pass Regressive Duration")]
            public double Total_First_Pass_Regressive_Duration
            {
                get
                {
                    if (m_Total_First_Pass_Regressive_Duration == -1)
                        m_Total_First_Pass_Regressive_Duration = this.First_Pass_Duration - this.Total_First_Pass_Progressive_Duration_Overall;
                    return m_Total_First_Pass_Regressive_Duration;
                }
            }

            private double m_Total_First_Pass_Regressive_Number = -1;
            [Description("Total First-Pass Regressive Number")]
            public double Total_First_Pass_Regressive_Number
            {
                get
                {
                    if (m_Total_First_Pass_Regressive_Number == -1)
                        m_Total_First_Pass_Regressive_Number = this.First_Pass_Number - this.Total_First_Pass_Progressive_Number_Overall;
                    return m_Total_First_Pass_Regressive_Number;
                }
            }

            private double m_Regression_Number = -1;
            [Description("Regression Number")]
            public double Regression_Number
            {
                get
                {
                    if (m_Regression_Number == -1)
                        m_Regression_Number = this.Regressions.Count();
                    return m_Regression_Number;
                }
            }

            private double m_Regression_Duration = -1;
            [Description("Regression Duration")]
            public double Regression_Duration
            {
                get
                {
                    if (m_Regression_Duration == -1)
                        m_Regression_Duration = this.Regressions.Sum(lst => lst.Sum(fix => fix.Event_Duration));
                    return m_Regression_Duration;
                }
            }

            private double m_First_Regression_Duration;
            [Description("First Regression Duration")]
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
                    return this.m_First_Regression_Duration;
                }
            }

            [Description("Skip")]
            public bool Skip { get; set; }

            private double m_Pupil_Diameter = -1;
            [Description("Pupil Diameter [mm]")]
            public double Pupil_Diameter
            {
                get
                {
                    if (m_Pupil_Diameter == -1)
                        m_Pupil_Diameter = this.Total_Pupil_Diameter / this.Total_Fixation_Number;
                    return m_Pupil_Diameter;
                }
            }

            private double m_Mean_AOI_Size = -1;
            [Description("AOI Size X [mm]")]
            public double Mean_AOI_Size
            {
                get
                {
                    if (m_Mean_AOI_Size == -1)
                    {
                        List<double> sizes = new List<double>();
                        IEnumerable<IEnumerable<double>> all_sizes = this.Fixations.Select(lst => lst.Select(fix=>{
                            if (currentType == AOITypes.Phrases)
                            {
                                if (fix.AOI_Name > 0 && !fix.IsInExceptionBounds)
                                    return fix.AOI_Phrase_Details.AOI_Size_X;
                                else
                                    return 0;
                            }
                            else if (currentType == AOITypes.Words)
                            {
                                if (fix.Word_Index > 0 && !fix.IsInExceptionBounds)
                                    return fix.AOI_Word_Details.AOI_Size_X;
                                else
                                    return 0;
                            }
                            return 0;
                        }).Distinct());
                        foreach (var lst in all_sizes)
                        {
                            sizes.AddRange(lst);
                        }
                        sizes = sizes.Distinct().ToList();
                        m_Mean_AOI_Size = sizes.Sum();
                    }
                    //m_Mean_AOI_Coverage = this.Total_AOI_Coverage / this.Total_Fixation_Number;
                    return m_Mean_AOI_Size;
                    //return this.Total_AOI_Size / this.Total_Fixation_Number;
                }
            }

            private double m_Mean_AOI_Coverage = -1;
            [Description("AOI Coverage [%]")]
            public double Mean_AOI_Coverage
            {
                get
                {
                    if (m_Mean_AOI_Coverage == -1)
                    {
                        List<double> coverages = new List<double>();
                        IEnumerable<IEnumerable<double>> all_coverages = this.Fixations.Select(lst => lst.Select(fix =>
                        {
                            if(currentType == AOITypes.Phrases)
                            {
                                if (fix.AOI_Name > 0 && !fix.IsInExceptionBounds)
                                    return fix.AOI_Phrase_Details.AOI_Coverage_In_Percents;
                                else
                                    return 0;
                            } else if (currentType == AOITypes.Words)
                            {
                                if (fix.Word_Index > 0 && !fix.IsInExceptionBounds)
                                    return fix.AOI_Word_Details.AOI_Coverage_In_Percents;
                                else
                                    return 0;
                            }
                            return 0;
                        }).Distinct());
                        foreach(var lst in all_coverages)
                        {
                            coverages.AddRange(lst);
                        }
                        coverages=coverages.Distinct().ToList();
                        m_Mean_AOI_Coverage = coverages.Sum();
                    }
                    return m_Mean_AOI_Coverage;
                }
            }
            
         

            private double m_Total_AOI_Size;
            [EpplusIgnore]
            public double Total_AOI_Size
            {
                get
                {
                    if (this.m_Total_AOI_Size == -1)
                        this.m_Total_AOI_Size = this.Fixations.Sum(lst => lst.Sum(fix => fix.AOI_Phrase_Size));
                    return this.m_Total_AOI_Size;
                }
            }

            private double m_Total_Pupil_Diameter;
            [EpplusIgnore]
            public double Total_Pupil_Diameter
            {
                get
                {
                    if (this.m_Total_Pupil_Diameter == -1)
                        this.m_Total_Pupil_Diameter = this.Fixations.Sum(lst => lst.Sum(fix => fix.Fixation_Average_Pupil_Diameter));
                    return this.m_Total_Pupil_Diameter;
                }
            }

            [EpplusIgnore]
            public List<List<Fixation>> Fixations { get; set; }

            [EpplusIgnore]
            public List<Fixation> First_Pass_Fixations { get; set; }

            // Length and Frequency should be calculated only for the AOI's by Words output, otherwise they are 0
            // and we delete them in the Excel itself
            public int Length
            {
                get
                {
                    if (currentType == AOITypes.Words && AOI_Group > 0)
                        return Fixations[0][0].AOI_Word_Details.Length;
                    return 0;
                }
            }
            public int Frequency
            {
                get
                {
                    if (currentType == AOITypes.Words && AOI_Group > 0)
                        return Fixations[0][0].AOI_Word_Details.Frequency;
                    return 0;
                }
            }

            private List<Fixation> m_Fixations_Progressive_First_Pass;
            private List<Fixation> Fixations_Progressive_First_Pass
            {
                get
                {
                    if (this.m_Fixations_Progressive_First_Pass == null)
                    {
                        if (this.First_Pass_Fixations.Any())
                        {
                            Fixation firstFixation = this.First_Pass_Fixations[0];
                            int firstOneRightToFirst = this.First_Pass_Fixations.FindIndex(fix => fix.Previous_Fixation != null && fix.IsBeforeThan(fix.Previous_Fixation));
                            if (firstOneRightToFirst == -1)
                                firstOneRightToFirst = this.First_Pass_Fixations.Count;
                            this.m_Fixations_Progressive_First_Pass = this.First_Pass_Fixations.GetRange(0, firstOneRightToFirst);
                        }
                        else
                        {
                            this.m_Fixations_Progressive_First_Pass = new List<Fixation>();
                        }
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
                        if (this.First_Pass_Fixations.Any())
                        {
                            Fixation[] first_Pass_Fixations = this.First_Pass_Fixations.ToArray();
                            this.m_Total_Fixations_Progressive_First_Pass.Add(first_Pass_Fixations[0]);
                            for (int i = 1 ; i < first_Pass_Fixations.Length ; ++i)
                            {
                                if (first_Pass_Fixations[i - 1].IsBeforeThan(first_Pass_Fixations[i]))
                                    this.m_Total_Fixations_Progressive_First_Pass.Add(first_Pass_Fixations[i]);
                            }
                        }
                    }
                    return this.m_Total_Fixations_Progressive_First_Pass;
                }
            }

            private List<List<Fixation>> m_Regressions;
            private IEnumerable<List<Fixation>> Regressions
            {
                get
                {
                    if (this.First_Pass_Fixations.Count == 0)
                        this.m_Regressions = new List<List<Fixation>>();
                    else if (this.m_Regressions == null)
                    {
                        this.m_Regressions = new List<List<Fixation>>();
                        this.Fixations.ForEach(lst =>
                        {
                            this.m_Regressions.Add(lst.ToList());
                        });
                        int firstPassIndex = -1;
                        int i = 0;
                        if (this.First_Pass_Fixations.Count > 0)
                        {
                            this.m_Regressions.ForEach((lst) =>
                            {
                                // FIND THE LAST FIXATION OF THE FIRST PASS, THE BATCHES OF THIS AOI FIXATIONS IS "NOT THE SAME"
                                // AS THE BATCHES FOR CALCULATION OF FIRST PASS
                                if (lst.Contains(this.First_Pass_Fixations[First_Pass_Fixations.Count - 1]))
                                {
                                    firstPassIndex = i;
                                }
                                i += 1;
                            });
                            i = 0;
                            m_Regressions.ForEach((lst) => {
                                if (i <= firstPassIndex)
                                    lst.Clear();
                                i += 1;
                            });

                        }
                        if (currentType == AOITypes.Phrases)
                            this.m_Regressions.RemoveAll(lst => !FixationsService.IsLeagalRegressionFixations(lst));
                        else if (currentType == AOITypes.Words)
                            this.m_Regressions.RemoveAll(lst => lst.Count == 0);

                    }
                    return this.m_Regressions;
                }
            }


            public AOIClass( string Stimulus, string Participant, int AOI_Group, bool Skip, AOITypes type)
            {
                this.Stimulus = Stimulus;
                this.Participant = Participant;
                this.AOI_Group = AOI_Group;
                if (type == AOITypes.Phrases)
                    this.AOI_Target = AOIDetails.groupPhraseToSpecialName.ContainsKey(AOI_Group) ? AOIDetails.groupPhraseToSpecialName[AOI_Group] : null;
                else
                    this.AOI_Target = AOIWordDetails.groupWordToSpecialName.ContainsKey(AOI_Group) ? AOIWordDetails.groupWordToSpecialName[AOI_Group] : null;
                if (this.AOI_Target != null)
                {
                    List<string> sNames = Constans.parseSpecialName(AOI_Target);
                    if (sNames.Count > maxConditions)
                        maxConditions = sNames.Count;
                }
                this.Skip = Skip;
                this.Fixations = new List<List<Fixation>>();
                
                this.m_Total_Fixation_Duration = -1;
                this.m_Total_Fixation_Number = -1;
                this.m_First_Pass_Duration = -1;
                this.m_First_Pass_Number = -1;
                this.m_Fixations_Progressive_First_Pass = null;
                this.First_Pass_Fixations = new List<Fixation>();
                this.m_Regressions = null;
                this.m_First_Pass_Progressive_Duration_Overall = -1;
                this.m_Total_First_Pass_Progressive_Duration_Overall = -1;
                this.m_First_Regression_Duration = -1;
                this.m_Total_Pupil_Diameter = -1;
                this.m_Total_AOI_Size = -1;
                new SecondFileConsideringCoverage.AIOClassAfterCoverage(this);

            }



        }
    }
}
