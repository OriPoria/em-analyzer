﻿using EM_Analyzer.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EM_Analyzer.ModelClasses
{
    class FixationsService
    {
        public static Dictionary<string, List<Fixation>> fixationSetToFixationListDictionary = new Dictionary<string, List<Fixation>>();
        public static string textFileName = "";
        public static string excelFileName = "";
        public static string outputPath = "";
        public static List<string> tableColumns;

        private class CountedAOIFixations
        {
            public List<Fixation> Fixations { get; set; }
            public int AOI_Group { get; set; }
            public int Count { get; set; }
        }

        private static List<CountedAOIFixations> ConvertFixationListToCoutedList(List<Fixation> fixations)
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

        public static void SortDictionary()
        {
            List<List<Fixation>> values = fixationSetToFixationListDictionary.Values.ToList();
            values.ForEach(fixationList => fixationList.Sort((a, b) => a.Index.CompareTo(b.Index)));
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

        public static void SearchForExceptions()
        {
            List<Fixation>[] values = fixationSetToFixationListDictionary.Values.ToArray();
            int Number_Of_Fixations_Out_Of_AOI_For_Exception = int.Parse(ConfigurationService.NumberOfFixationsOutOfAOIForException);
            int Number_Of_Fixations_In_AOI_For_Exception = int.Parse(ConfigurationService.NumberOfFixationsInAOIForException);
            double exceptionsLimit = double.Parse(ConfigurationService.DealingWithExceptionsLimitInPixels);
            Queue<Fixation> lastFixationsQueue;// = new Queue<Fixation>();
            while (Number_Of_Fixations_In_AOI_For_Exception > 0)
            {
                lastFixationsQueue = new Queue<Fixation>();
                foreach (List<Fixation> fixationList in values)
                {
                    lastFixationsQueue.Clear();
                    foreach (Fixation fixation in fixationList)
                    {
                        // Get the last Fixations that relevant for our window
                        if (lastFixationsQueue.Count == Number_Of_Fixations_Out_Of_AOI_For_Exception + Number_Of_Fixations_In_AOI_For_Exception - 1)
                        {
                            if (fixation.AOI_Group_After_Change == lastFixationsQueue.First().AOI_Group_After_Change)
                            {
                                // Gets a list of all the fixations that have another AOI.
                                List<Fixation> notAOIEqualsFixations = lastFixationsQueue.Where(fix => fix.AOI_Group_After_Change != fixation.AOI_Group_After_Change).ToList();

                                if (notAOIEqualsFixations.Count <= Number_Of_Fixations_In_AOI_For_Exception) // it is exception
                                {
                                    notAOIEqualsFixations.ForEach(exceptionalFix =>
                                    {
                                        exceptionalFix.IsException = true;
                                        exceptionalFix.AOI_Group_After_Change = fixation.AOI_Group_After_Change;
                                        // fix.IsInExceptionBounds = (fix.DistanceToAOI(AOIDetails.nameToAOIDetailsDictionary[fixation.AOI_Name+ fixation.Stimulus]) <= exceptionsLimit);
                                        exceptionalFix.IsInExceptionBounds = fixation.AOI_Name != -1 && (exceptionalFix.DistanceToAOI(fixation.AOI_Details) <= exceptionsLimit);
                                    });
                                }
                            }
                            lastFixationsQueue.Dequeue();
                        }
                        lastFixationsQueue.Enqueue(fixation);
                    }
                }
                Number_Of_Fixations_In_AOI_For_Exception--;
            }

            Number_Of_Fixations_In_AOI_For_Exception = int.Parse(ConfigurationService.NumberOfFixationsInAOIForException);
            foreach (List<Fixation> fixationList in values)
            {
                CountedAOIFixations[] countedAOIFixationsArray = ConvertFixationListToCoutedList(fixationList).ToArray();
                for (int i = 0; i < countedAOIFixationsArray.Length; i++)
                {
                    CountedAOIFixations currrentCountedAOIFixations = countedAOIFixationsArray[i];
                    if (currrentCountedAOIFixations.Count <= Number_Of_Fixations_In_AOI_For_Exception)
                    {
                        if (i > 0 && countedAOIFixationsArray[i - 1].Count >= Number_Of_Fixations_Out_Of_AOI_For_Exception)
                        {
                            currrentCountedAOIFixations.Fixations.ForEach(fix => fix.AOI_Group_After_Change = countedAOIFixationsArray[i - 1].AOI_Group);
                            countedAOIFixationsArray[i - 1].Count += currrentCountedAOIFixations.Count;
                            countedAOIFixationsArray[i - 1].Fixations.AddRange(currrentCountedAOIFixations.Fixations);
                            countedAOIFixationsArray[i] = countedAOIFixationsArray[i - 1];
                        }
                        else
                        {
                            if (i < countedAOIFixationsArray.Length-1 && countedAOIFixationsArray[i + 1].Count >= Number_Of_Fixations_Out_Of_AOI_For_Exception)
                            {
                                currrentCountedAOIFixations.Fixations.ForEach(fix => fix.AOI_Group_After_Change = countedAOIFixationsArray[i + 1].AOI_Group);
                                countedAOIFixationsArray[i + 1].Count += currrentCountedAOIFixations.Count;
                                countedAOIFixationsArray[i + 1].Fixations.AddRange(currrentCountedAOIFixations.Fixations);
                                countedAOIFixationsArray[i] = countedAOIFixationsArray[i + 1];
                            }
                        }
                    }
                }
            }
        }

        public static void DealWithExceptions()
        {
            List<Fixation>[] values = fixationSetToFixationListDictionary.Values.ToArray();
            DealingWithExceptionsEnum dealingWithInsideExceptions = (DealingWithExceptionsEnum)int.Parse(ConfigurationService.DealingWithExceptionsInsideTheLimit);
            DealingWithExceptionsOutBoundsEnum dealingWithOutsideExceptions = (DealingWithExceptionsOutBoundsEnum)int.Parse(ConfigurationService.DealingWithExceptionsOutsideTheLimit);
            foreach (List<Fixation> fixationList in values)
            {
                IEnumerable<Fixation> exceptionsFixations = fixationList.Where(fix => fix.IsException);
                foreach (Fixation fixation in exceptionsFixations)
                {
                    if (fixation.IsInExceptionBounds && dealingWithInsideExceptions > DealingWithExceptionsEnum.Do_Nothing)
                    {
                        fixation.AOI_Group_Before_Change = fixation.AOI_Group_After_Change;
                        fixation.IsException = false;
                    }
                }
            }
        }

        // For the text file with the bad headers (not relevant for now, maybe for the future).
        public static void InitializeColumnIndexes()
        {
            TextFileColumnIndexes.Trial = tableColumns.FindIndex(column => column.Trim().ToLower().StartsWith("trial"));
            TextFileColumnIndexes.Stimulus = tableColumns.FindIndex(column => column.Trim().ToLower().StartsWith("stimulus"));
            TextFileColumnIndexes.Participant = tableColumns.FindIndex(column => column.Trim().ToLower().StartsWith("participant"));
            TextFileColumnIndexes.AOI_Name = tableColumns.FindIndex(column => column.Trim().ToLower().StartsWith("aoiname"));
            TextFileColumnIndexes.AOI_Group = tableColumns.FindIndex(column => column.Trim().ToLower().StartsWith("aoigroup"));
            TextFileColumnIndexes.AOI_Coverage = tableColumns.FindIndex(column => column.Trim().ToLower().StartsWith("aoicoverage"));
            TextFileColumnIndexes.Index = tableColumns.FindIndex(column => column.Trim().ToLower().StartsWith("index"));
            TextFileColumnIndexes.Event_Duration = tableColumns.FindIndex(column => column.Trim().ToLower().StartsWith("event"));
            TextFileColumnIndexes.Fixation_Position_Y = tableColumns.FindIndex(column => column.Trim().ToLower().Contains("position") && column.Trim().ToLower().Contains("y"));
            TextFileColumnIndexes.Fixation_Position_X = tableColumns.FindIndex(column => column.Trim().ToLower().Contains("position") && !column.Trim().ToLower().Contains("y"));
            TextFileColumnIndexes.Fixation_Average_Pupil_Diameter = tableColumns.FindIndex(column => column.Trim().ToLower().Contains("average"));
        }
    }
}
