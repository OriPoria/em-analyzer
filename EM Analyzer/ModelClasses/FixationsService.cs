﻿using EM_Analyzer.Enums;
using EM_Analyzer.ModelClasses.AOIClasses;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using EM_Analyzer.ExcelLogger;
using System.Windows.Forms;

namespace EM_Analyzer.ModelClasses
{

    public class CountedAOIFixations
    {
        public List<Fixation> Fixations { get; set; }
        public int AOI_Group { get; set; }
        public int Count { get; set; }
    }
    class FixationsService
    {
        public static Dictionary<string, List<Fixation>> fixationSetToFixationListDictionary = new Dictionary<string, List<Fixation>>();
        public static Dictionary<string, List<WordIndex>> wordIndexSetToFixationListDictionary = new Dictionary<string, List<WordIndex>>();
        public static string phrasesTextFileName = "";
        public static string wordsTextFileName = "";
        public static string phrasesExcelFileName = "";
        public static string wordsExcelFileName = "";
        public static string outputPath = "";
        public static List<string> tableColumns;
        public static DealingWithExceptionsEnum dealingWithInsideExceptions = (DealingWithExceptionsEnum)int.Parse(ConfigurationService.DealingWithExceptionsInsideTheLimit);
        public static DealingWithExceptionsOutBoundsEnum dealingWithOutsideExceptions = (DealingWithExceptionsOutBoundsEnum)int.Parse(ConfigurationService.DealingWithExceptionsOutsideTheLimit);
        
        public static int Number_Of_Fixations_In_Of_AOI_For_Exception;
        public static int Number_Of_Fixations_Out_AOI_For_Exception;
        
        public static int Minimum_Number_Of_Fixations_For_First_Pass = int.Parse(ConfigurationService.MinimumNumberOfFixationsForFirstPass);
        public static double Minimum_Duration_Of_Fixation_For_First_Pass = double.Parse(ConfigurationService.MinimumDurationOfFixationForFirstPass);
        
        public static int Minimum_Number_Of_Fixations_For_Skip = int.Parse(ConfigurationService.MinimumNumberOfFixationsForSkip);
        public static double Minimum_Duration_Of_Fixation_For_Skip = double.Parse(ConfigurationService.MinimumDurationOfFixationForSkip);
        
        public static int Minimum_Number_Of_Fixations_For_Regression = int.Parse(ConfigurationService.MinimumNumberOfFixationsForRegression);
        public static double Minimum_Duration_Of_Fixation_For_Regression = double.Parse(ConfigurationService.MinimumDurationOfFixationForRegression);

        public static double exceptionsLimit = double.Parse(ConfigurationService.DealingWithExceptionsLimitInPixels);

        //public static bool IsFixationShouldBeSkippedInFirstPass(Fixation)

        public static List<CountedAOIFixations> ConvertFixationListToCoutedListByPhrase(List<Fixation> fixations)
        {
            List<CountedAOIFixations> countedAOIFixations = new List<CountedAOIFixations>();
            Fixation prevFixation = fixations.First();
            CountedAOIFixations currentCountedAOIFixations = new CountedAOIFixations() { AOI_Group = fixations.First().AOI_Group_After_Change, Count = 0, Fixations = new List<Fixation>() };
            countedAOIFixations.Add(currentCountedAOIFixations);
            foreach (Fixation fixation in fixations)
            {
                if (fixation.AOI_Group_After_Change == prevFixation.AOI_Group_After_Change)
                {
                    currentCountedAOIFixations.Count++;
                }
                else
                {
                    currentCountedAOIFixations = new CountedAOIFixations() { AOI_Group = fixation.AOI_Group_After_Change, Count = 1, Fixations = new List<Fixation>() };
                    countedAOIFixations.Add(currentCountedAOIFixations);
                }
                currentCountedAOIFixations.Fixations.Add(fixation);
                prevFixation = fixation;
            }
            return countedAOIFixations;
        }
        public static List<CountedAOIFixations> ConvertFixationListToCoutedListByWords(List<Fixation> fixations)
        {
            List<CountedAOIFixations> countedAOIFixations = new List<CountedAOIFixations>();
            Fixation prevFixation = fixations.First();
            CountedAOIFixations currentCountedAOIFixations = new CountedAOIFixations() { AOI_Group = fixations.First().Word_Index, Count = 0, Fixations = new List<Fixation>() };
            countedAOIFixations.Add(currentCountedAOIFixations);
            foreach (Fixation fixation in fixations)
            {
                if (fixation.Word_Index == prevFixation.Word_Index)
                {
                    currentCountedAOIFixations.Count++;
                }
                else
                {
                    currentCountedAOIFixations = new CountedAOIFixations() { AOI_Group = fixation.Word_Index, Count = 1, Fixations = new List<Fixation>() };
                    countedAOIFixations.Add(currentCountedAOIFixations);
                }
                currentCountedAOIFixations.Fixations.Add(fixation);
                prevFixation = fixation;
            }
            return countedAOIFixations;
        }
        public static void SortDictionary()
        {
            List<List<Fixation>> values = fixationSetToFixationListDictionary.Values.ToList();
            values.ForEach(fixationList => fixationList.Sort((a, b) => a.Index.CompareTo(b.Index)));
        }
        public static void SortWordIndexDictionary()
        {
            List<List<WordIndex>> values = wordIndexSetToFixationListDictionary.Values.ToList();
            values.ForEach(fixationList => fixationList.Sort((a, b) => a.Index.CompareTo(b.Index)));
        }
        public static void UnifyDictionaryWithWordIndex()
        {
            foreach (var item in fixationSetToFixationListDictionary.Keys)
            {
                var x1 = fixationSetToFixationListDictionary[item];
                var x2 = wordIndexSetToFixationListDictionary[item];
                for (int i = 0; i < x1.Count; i++)
                {
                    if (x1[i].Index != x2[i].Index)
                        MessageBox.Show($"error in sort of fixation by index at line {i} at key {item}");
                    x1[i].Word_Index = x2[i].Group;
                    x1[i].AOI_Word_Size = x2[i].AOI_Word_Size;
                    if (x1[i].Word_Index != -1 && x1[i].AOI_Word_Details.AOI_Size_X < 0)
                    {
                        x1[i].AOI_Word_Details.AOI_Size_X = x1[i].AOI_Word_Size;
                    }
                }
            }
        }

        public static void CleanAllFixationBeforeFirstAOI()
        {
            List<Fixation>[] values = fixationSetToFixationListDictionary.Values.ToArray();
            foreach (List<Fixation> fixationList in values)
            {

                int firstFixaitionAtFirstAOI = fixationList.FindIndex(fix => fix.AOI_Group_Before_Change == 1);
                if (firstFixaitionAtFirstAOI > 0)
                    fixationList.RemoveRange(0, firstFixaitionAtFirstAOI);
            }
        }
        public static bool IsLeagalFirstPassFixations(CountedAOIFixations countedAOIFixations)
        {
            int counter = 0;
            foreach (Fixation fixation in countedAOIFixations.Fixations)
            {
                if (fixation.Event_Duration > Minimum_Duration_Of_Fixation_For_First_Pass)
                    counter += 1;

                if (counter >= Minimum_Number_Of_Fixations_For_First_Pass)
                    return true;
            }

            return false;
        }
        public static bool IsLeagalFixationsForSkip(List<Fixation> fixationRange)
        {
            int counter = 0;
            foreach (Fixation fixation in fixationRange)
            {
                if (fixation.Event_Duration > Minimum_Duration_Of_Fixation_For_Skip)
                    counter += 1;
                if (counter >= Minimum_Number_Of_Fixations_For_Skip)
                    return true;
            }
            return false;
        }
        public static bool IsLeagalRegressionFixations(List<Fixation> fixations) 
        {
            int counter = 0;
            foreach (Fixation fixation in fixations)
            {
                if (fixation.Event_Duration > Minimum_Duration_Of_Fixation_For_Regression)
                    counter += 1;
                if (counter >= Minimum_Number_Of_Fixations_For_Regression)
                    return true;
            }
            return false;
        }

        public static void DealWithSeparatedAOIs()
        {
            List<AOIDetails> all_AOIs = AOIDetails.nameToAOIPhrasesDetailsDictionary.Values.ToList();

            var results = from aoi in all_AOIs
                          group aoi by aoi.Stimulus + aoi.Group into grou
                          select new { Key = grou.Key, AIOS = grou.ToList() };

            Dictionary<string, List<AOIDetails>> dict = new Dictionary<string, List<AOIDetails>>();
            results.ToList().ForEach(tuple => dict[tuple.Key] = tuple.AIOS);

            IEnumerable<List<AOIDetails>> separatedAOIs = dict.Values.Where(lst => lst.Count > 1); ;
            foreach (var AOIs in separatedAOIs)
            {
                SeparatedAOI separatedAOI = new SeparatedAOI(AOIs.Cast<IAOI>());
                foreach (var aoi in AOIs)
                {
                    AOIsService.nameToAOIPhrasesDictionary[aoi.DictionaryKey] = separatedAOI;
                }
            }
        }

        public static void SearchForExceptions()
        {
            List<Fixation>[] values = fixationSetToFixationListDictionary.Values.ToArray(); // all fixations per participent 
            Number_Of_Fixations_In_Of_AOI_For_Exception = int.Parse(ConfigurationService.NumberOfFixationsInOfAOIForException);
            Number_Of_Fixations_Out_AOI_For_Exception = int.Parse(ConfigurationService.NumberOfFixationsOutAOIForException);
            //double exceptionsLimit = double.Parse(ConfigurationService.DealingWithExceptionsLimitInPixels);
            Queue<Fixation> lastFixationsQueue;// = new Queue<Fixation>();
            while (Number_Of_Fixations_Out_AOI_For_Exception > 0)
            {
                lastFixationsQueue = new Queue<Fixation>();
                foreach (List<Fixation> fixationList in values)
                {
                    lastFixationsQueue.Clear();
                    List<Fixation> notExceptionalFixations = fixationList.Where(fix => !fix.IsException || (fix.IsInExceptionBounds && dealingWithInsideExceptions == DealingWithExceptionsEnum.Change_AOI_Group)).ToList();
                    foreach (Fixation fixation in notExceptionalFixations)
                    {

                        // Get the last Fixations that relevant for our window
                        if (lastFixationsQueue.Count == Number_Of_Fixations_In_Of_AOI_For_Exception + Number_Of_Fixations_Out_AOI_For_Exception - 1)
                        {
                            if (fixation.AOI_Group_After_Change == lastFixationsQueue.First().AOI_Group_After_Change)
                            {
                                // Gets a list of all the fixations that have another AOI in the last fixations queue 
                                List<Fixation> notAOIEqualsFixations = lastFixationsQueue.Where(fix => fix.AOI_Group_After_Change != fixation.AOI_Group_After_Change).ToList();

                                if (notAOIEqualsFixations.Count <= Number_Of_Fixations_Out_AOI_For_Exception && notAOIEqualsFixations.Any()) // it is exception
                                {
                                    notAOIEqualsFixations.ForEach(exceptionalFix =>
                                    {
                                        exceptionalFix.IsException = true;
                                        exceptionalFix.IsInExceptionBounds = fixation.AOI_Name != -1 && (fixation.AOI_Phrase_Details.DistanceToAOI(exceptionalFix) <= exceptionsLimit);
                                        if (exceptionalFix.IsInExceptionBounds && dealingWithInsideExceptions == DealingWithExceptionsEnum.Change_AOI_Group)
                                            exceptionalFix.AOI_Group_After_Change = fixation.AOI_Group_After_Change;
                                        // fix.IsInExceptionBounds = (fix.DistanceToAOI(AOIDetails.nameToAOIDetailsDictionary[fixation.AOI_Name+ fixation.Stimulus]) <= exceptionsLimit);
                                    });
                                    lastFixationsQueue = new Queue<Fixation>(lastFixationsQueue.Where(fix => !fix.IsException));
                                }
                            }
                            if(lastFixationsQueue.Any())
                                lastFixationsQueue.Dequeue();
                        }
                        lastFixationsQueue.Enqueue(fixation);
                    }
                }
                Number_Of_Fixations_Out_AOI_For_Exception--;
            }

            Number_Of_Fixations_Out_AOI_For_Exception = int.Parse(ConfigurationService.NumberOfFixationsOutAOIForException);
            foreach (List<Fixation> fixationList in values)
            {
                List<Fixation> notExceptionalFixations = fixationList.Where(fix => !fix.IsException || (fix.IsInExceptionBounds && dealingWithInsideExceptions == DealingWithExceptionsEnum.Change_AOI_Group)).ToList();
                //fixationList.RemoveAll(fix => fix.IsException);
                List<CountedAOIFixations> countedAOIFixationsArray = ConvertFixationListToCoutedListByPhrase(notExceptionalFixations).ToList();
                for (int i = 0 ; i < countedAOIFixationsArray.Count ; i++)
                {
                    CountedAOIFixations currrentCountedAOIFixations = countedAOIFixationsArray[i];
                    if (i > 0 && currrentCountedAOIFixations.AOI_Group == countedAOIFixationsArray[i - 1].AOI_Group)
                    {
                        countedAOIFixationsArray[i - 1].Count += currrentCountedAOIFixations.Count;
                        countedAOIFixationsArray[i - 1].Fixations.AddRange(currrentCountedAOIFixations.Fixations);
                        countedAOIFixationsArray.RemoveAt(i);
                        i--;
                    }
                    else
                    {
                        if (currrentCountedAOIFixations.Count <= Number_Of_Fixations_Out_AOI_For_Exception)
                        {
                            i = ChangeAOIGroupOfCountedAOIFixations(countedAOIFixationsArray, i);
                         
                        }
                    }
                }
            }
        }

        private static void AddCountedAOIToAnother(List<CountedAOIFixations> countedAOIFixationsArray, int FromIndex, int ToIndex)
        {
            IAOI aoiAddingTo = countedAOIFixationsArray[ToIndex].Fixations.First().AOI_Phrase_Details;
            countedAOIFixationsArray[FromIndex].Fixations.ForEach(fix =>
            {
                fix.IsInExceptionBounds = countedAOIFixationsArray[ToIndex].Fixations.First().AOI_Name != -1 && (aoiAddingTo.DistanceToAOI(fix) <= exceptionsLimit);
                if (fix.IsInExceptionBounds && dealingWithInsideExceptions == DealingWithExceptionsEnum.Change_AOI_Group)
                    fix.AOI_Group_After_Change = countedAOIFixationsArray[ToIndex].AOI_Group;
            });
            countedAOIFixationsArray[ToIndex].Count += countedAOIFixationsArray[FromIndex].Count;
            countedAOIFixationsArray[ToIndex].Fixations.AddRange(countedAOIFixationsArray[FromIndex].Fixations);
            countedAOIFixationsArray.RemoveAt(FromIndex);
        }

        private static int ChangeAOIGroupOfCountedAOIFixations(List<CountedAOIFixations> countedAOIFixationsArray, int index)
        {
            CountedAOIFixations currrentCountedAOIFixations = countedAOIFixationsArray[index];
            bool needsToAddToPrevAOI = index > 0 && countedAOIFixationsArray[index - 1].Count >= Number_Of_Fixations_In_Of_AOI_For_Exception && countedAOIFixationsArray[index - 1].AOI_Group!=-1;
            bool needsToAddToNextAOI = index < countedAOIFixationsArray.Count - 1 && countedAOIFixationsArray[index + 1].Count >= Number_Of_Fixations_In_Of_AOI_For_Exception && countedAOIFixationsArray[index + 1].AOI_Group != -1;
            if (needsToAddToPrevAOI || needsToAddToNextAOI)
                currrentCountedAOIFixations.Fixations.ForEach(fix => fix.IsException = true);
            // TODO: Dealing With Exceptional limits !!!!!!!!!!!!!!!!!!!!
            if (needsToAddToPrevAOI && needsToAddToNextAOI)
            {
                CountedAOIFixations prev = countedAOIFixationsArray[index - 1], next = countedAOIFixationsArray[index + 1];
                int firstClosedToNext = currrentCountedAOIFixations.Fixations.FindIndex(fixation => prev.Fixations.First().AOI_Phrase_Details.DistanceToAOI(fixation) >= next.Fixations.First().AOI_Phrase_Details.DistanceToAOI(fixation));

                if (firstClosedToNext < 0)
                {
                    IAOI aoiAddingTo = prev.Fixations.First().AOI_Phrase_Details;
                    currrentCountedAOIFixations.Fixations.ForEach(fix =>
                    {
                        
                        fix.IsInExceptionBounds = prev.Fixations.First().AOI_Name != -1 && (aoiAddingTo.DistanceToAOI(fix) <= exceptionsLimit);
                        if (fix.IsInExceptionBounds && dealingWithInsideExceptions == DealingWithExceptionsEnum.Change_AOI_Group)
                            fix.AOI_Group_After_Change = prev.AOI_Group;
                            //ExcelLoggerService.AddLog(log: new Log() { FileName = "Hanaka_reduced - First", LineNumber = (uint)fix.Index, Description = "The Fixation Is Not Inside The AOI " + fix.AOI_Group_Before_Change });
                    });
                    prev.Count += currrentCountedAOIFixations.Count;
                    prev.Fixations.AddRange(currrentCountedAOIFixations.Fixations);
                }
                else
                {
                    IEnumerable<Fixation> fixationsAddToNextAOI = currrentCountedAOIFixations.Fixations.GetRange(firstClosedToNext, currrentCountedAOIFixations.Fixations.Count - firstClosedToNext);
                    IAOI aoiAddingTo = next.Fixations.First().AOI_Phrase_Details;
                    foreach (var fix in fixationsAddToNextAOI)
                    {
                        fix.IsInExceptionBounds = next.Fixations.First().AOI_Name != -1 && (aoiAddingTo.DistanceToAOI(fix) <= exceptionsLimit);
                        if (fix.IsInExceptionBounds && dealingWithInsideExceptions == DealingWithExceptionsEnum.Change_AOI_Group)
                            fix.AOI_Group_After_Change = next.AOI_Group;
                    }
                    next.Count += fixationsAddToNextAOI.Count();
                    next.Fixations.AddRange(fixationsAddToNextAOI);

                    IEnumerable<Fixation> fixationsAddToPrevAOI = currrentCountedAOIFixations.Fixations.GetRange(0, firstClosedToNext);
                    aoiAddingTo = prev.Fixations.First().AOI_Phrase_Details;
                    foreach (var fix in fixationsAddToPrevAOI)
                    {
                        fix.IsInExceptionBounds = prev.Fixations.First().AOI_Name != -1 && (aoiAddingTo.DistanceToAOI(fix) <= exceptionsLimit);
                        if (fix.IsInExceptionBounds && dealingWithInsideExceptions == DealingWithExceptionsEnum.Change_AOI_Group)
                            fix.AOI_Group_After_Change = prev.AOI_Group;
                    }
                    prev.Count += fixationsAddToPrevAOI.Count();
                    prev.Fixations.AddRange(fixationsAddToPrevAOI);
                }
                countedAOIFixationsArray.Remove(currrentCountedAOIFixations);
            }
            else
            {
                if (needsToAddToPrevAOI)
                {
                    AddCountedAOIToAnother(countedAOIFixationsArray, index, index - 1);
                    index--;
                }
                else
                {
                    if (needsToAddToNextAOI)
                    {
                        AddCountedAOIToAnother(countedAOIFixationsArray, index, index + 1);
                        //index++;
                    }
                }
            }
            return index;
        }

        public static void DealWithExceptions()
        {
            List<Fixation>[] values = fixationSetToFixationListDictionary.Values.ToArray();
            foreach (List<Fixation> fixationList in values)
            {
                IEnumerable<Fixation> exceptionsFixations = fixationList.Where(fix => fix.IsException);
                foreach (Fixation fixation in exceptionsFixations)
                {
                    if (fixation.IsInExceptionBounds)
                    {
                        switch (dealingWithInsideExceptions)
                        {
                            case DealingWithExceptionsEnum.Do_Nothing:
                            case DealingWithExceptionsEnum.Skip_In_First_Pass:
                                fixation.AOI_Group_After_Change = fixation.AOI_Group_Before_Change;
                                break;
                            case DealingWithExceptionsEnum.Change_AOI_Group:
                                fixation.AOI_Group_Before_Change = fixation.AOI_Group_After_Change;
                                fixation.IsException = false;
                                break;
                        }
                    }
                }
            }
        }
        public static void ClearPreviousFixation(List<Fixation> fixations)
        {
            foreach (Fixation fixation in fixations)
            {
                fixation.Previous_Fixation = null;
            }
        }


        // For the text file with the bad headers (not relevant for now, maybe for the future).
        public static void InitializeColumnIndexes()
        {
            TextFileColumnIndexes.Trial = tableColumns.FindIndex(column => column.Equals("trial"));
            TextFileColumnIndexes.Stimulus = tableColumns.FindIndex(column => column.Equals("stimulus"));
            TextFileColumnIndexes.Participant = tableColumns.FindIndex(column => column.Equals("participant"));
            TextFileColumnIndexes.AOI_Name = tableColumns.FindIndex(column => column.StartsWith("aoi") && column.Contains("name"));
            TextFileColumnIndexes.AOI_Group = tableColumns.FindIndex(column => column.StartsWith("aoi") && column.Contains("group"));
            TextFileColumnIndexes.AOI_Size = tableColumns.FindIndex(column => column.StartsWith("aoi") && column.Contains("size"));
            TextFileColumnIndexes.AOI_Coverage = tableColumns.FindIndex(column => column.StartsWith("aoi") && column.Contains("coverage"));
            TextFileColumnIndexes.Index = tableColumns.FindIndex(column => column.Equals("index"));
            TextFileColumnIndexes.Event_Duration = tableColumns.FindIndex(column => column.StartsWith("event") && column.Contains("duration"));
            TextFileColumnIndexes.Fixation_Position_Y = tableColumns.FindIndex(column => column.Contains("position") && column.Contains("y"));
            TextFileColumnIndexes.Fixation_Position_X = tableColumns.FindIndex(column => column.Contains("position") && !column.Contains("y"));
            TextFileColumnIndexes.Fixation_Average_Pupil_Diameter = tableColumns.FindIndex(column => column.Contains("average") && column.Contains("pupil") && column.Contains("diameter"));
            // we set word index as the last column, not found in the first line of the titles of the column
            TextFileColumnIndexes.Word_Index = tableColumns.Count;
        }
    }
}
