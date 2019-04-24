using EM_Analyzer.Enums;
using EM_Analyzer.ExcelLogger;
using EM_Analyzer.ModelClasses.AOIClasses;
using EM_Analyzer.Services;

using System;
using System.Collections.Generic;
using System.ComponentModel;


namespace EM_Analyzer.ModelClasses
{
    public class Fixation
    {
        [Description("Trial")]
        public string Trial { get; set; }
        [Description("Stimulus")]
        public string Stimulus { get; set; }
        //[XLColumn(Ignore = true)]
        [EpplusIgnore]
        public Fixation Previous_Fixation { get; set; }
        [Description("Participant")]
        public string Participant { get; set; }
        [Description("AOI Name")]
        public int AOI_Name { get; set; }
        [Description("AOI Group Before Change")]
        public int AOI_Group_Before_Change { get; set; }
        [Description("AOI Group After Change")]
        public int AOI_Group_After_Change { get; set; }
        [Description("Is Exceptional")]
        public bool IsException { get; set; }
        [EpplusIgnore]
        public long AOI_Size { get; set; }
        [Description("AOI Coverage In Percents")]
        public double AOI_Coverage_In_Percents { get; set; }
        [Description("Index")]
        public long Index { get; set; }
        [Description("Event Duration")]
        public double Event_Duration { get; set; }
        [Description("Position X")]
        public double Fixation_Position_X { get; set; }
        [Description("Position Y")]
        public double Fixation_Position_Y { get; set; }
        [Description("Average Pupil Diameter")]
        public double Fixation_Average_Pupil_Diameter { get; set; }
        [EpplusIgnore]
        public IAOI AOI_Details
        {
            get
            {
                if (AOIDetails.isAOIIncludeStimulus)
                    return AOIsService.nameToAOIDictionary[AOI_Name + Stimulus];
                else
                    return AOIsService.nameToAOIDictionary[AOI_Name + ""];
            }
        }

        //[EpplusIgnore]
        [Description("Is In Exception Bounds")]
        public bool IsInExceptionBounds { get; set; }

        //static double minimumDuration = double.Parse(ConfigurationService.getValue(ConfigurationService.Minimum_Event_Duration_In_ms));
        //static double maximumDuration = double.Parse(ConfigurationService.getValue(ConfigurationService.Maximum_Event_Duration_In_ms));

        static readonly double minimumDuration = double.Parse(ConfigurationService.MinimumEventDurationInms);
        static readonly double maximumDuration = double.Parse(ConfigurationService.MaximumEventDurationInms);


        public static Fixation CreateFixationFromStringArray(string[] arr, uint lineNumber)
        {
            bool isFixationValid = true;
            bool isAOIValid = true;
            Fixation newFixation = new Fixation
            {
                Trial = arr[TextFileColumnIndexes.Trial].Trim(),
                Stimulus = arr[TextFileColumnIndexes.Stimulus].Trim(),
                Participant = arr[TextFileColumnIndexes.Participant].Trim()
            };

            try
            {
                newFixation.Event_Duration = double.Parse(arr[TextFileColumnIndexes.Event_Duration]);
                if (newFixation.Event_Duration < minimumDuration || newFixation.Event_Duration > maximumDuration)
                    return null;
            }

            catch
            {
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("Event Duration", arr[TextFileColumnIndexes.Event_Duration], lineNumber));
                isFixationValid = false;
            }

            try
            {
                newFixation.AOI_Group_After_Change = newFixation.AOI_Group_Before_Change = int.Parse(arr[TextFileColumnIndexes.AOI_Group]);
            }
            catch
            {
                isAOIValid = false;
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("AOI Group", arr[TextFileColumnIndexes.AOI_Group], lineNumber));
            }

            try
            {
                newFixation.AOI_Name = int.Parse(arr[TextFileColumnIndexes.AOI_Name]);
            }
            catch
            {
                isAOIValid = false;
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("AOI Name", arr[TextFileColumnIndexes.AOI_Name], lineNumber));
            }

            try
            {
                newFixation.AOI_Size = long.Parse(arr[TextFileColumnIndexes.AOI_Size]);
            }
            catch
            {
                isAOIValid = false;
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("AOI Size", arr[TextFileColumnIndexes.AOI_Size], lineNumber));
            }

            if (!isAOIValid)
            {
                newFixation.AOI_Group_After_Change = newFixation.AOI_Group_Before_Change = -1;
                newFixation.AOI_Name = -1;
                newFixation.AOI_Size = -1;
            }


            try
            {
                newFixation.Fixation_Position_X = double.Parse(arr[TextFileColumnIndexes.Fixation_Position_X]);
            }
            catch
            {
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("Fixation Position X", arr[TextFileColumnIndexes.Fixation_Position_X], lineNumber));
                isFixationValid = false;
            }

            try
            {
                newFixation.Fixation_Position_Y = double.Parse(arr[TextFileColumnIndexes.Fixation_Position_Y]);
            }
            catch
            {
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("Fixation Position Y", arr[TextFileColumnIndexes.Fixation_Position_Y], lineNumber));
                isFixationValid = false;
            }

            try
            {
                newFixation.Fixation_Average_Pupil_Diameter = double.Parse(arr[TextFileColumnIndexes.Fixation_Average_Pupil_Diameter]);
            }
            catch
            {
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("Fixation Average Pupil Diameter", arr[TextFileColumnIndexes.Fixation_Average_Pupil_Diameter], lineNumber));
                isFixationValid = false;
            }

            try
            {
                newFixation.AOI_Coverage_In_Percents = double.Parse(arr[TextFileColumnIndexes.AOI_Coverage]);
                if (newFixation.AOI_Name != -1 && newFixation.AOI_Details.AOI_Coverage_In_Percents < 0)
                {
                    newFixation.AOI_Details.AOI_Coverage_In_Percents = newFixation.AOI_Coverage_In_Percents;

                }
            }
            catch
            {
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("AOI_Coverage", arr[TextFileColumnIndexes.AOI_Coverage], lineNumber));
                isFixationValid = false;
            }

            try
            {
                newFixation.Index = long.Parse(arr[TextFileColumnIndexes.Index]);
            }
            catch
            {
                ExcelLoggerService.AddLog(CreateLogForFieldValidation("Index", arr[TextFileColumnIndexes.Index], lineNumber));
                isFixationValid = false;
            }



            newFixation.IsException = false;

            if (newFixation.AOI_Name != -1)
            {
                //if (AOIDetails.isAOIIncludeStimulus)
                //    newFixation.AOI_Details = AOIDetails.nameToAOIDetailsDictionary[newFixation.AOI_Name + newFixation.Stimulus];
                //else
                //    newFixation.AOI_Details = AOIDetails.nameToAOIDetailsDictionary[newFixation.AOI_Name + ""];

                //if (newFixation.AOI_Details.AOI_Coverage_In_Percents < 0)
                //{
                //    newFixation.AOI_Details.AOI_Coverage_In_Percents = newFixation.AOI_Coverage_In_Percents;
                //}

                //if (newFixation.AOI_Details.IsProper && newFixation.AOI_Details.DistanceToAOI(newFixation) != 0)
                if (newFixation.AOI_Details.DistanceToAOI(newFixation) != 0)
                {
                    //newFixation.AOI_Details.IsProper = false;
                    ExcelLoggerService.AddLog(new Log() { FileName = FixationsService.textFileName, LineNumber = lineNumber, Description = "The Fixation Is Not Inside The AOI " + newFixation.AOI_Name });
                    //new Task(() => MessageBox.Show("There is a problem with the AOI " + newFixation.AOI_Name + " In Stimulus " + newFixation.Stimulus)).Start();
                }
            }

            if (!isFixationValid)
                return null;

            string dictionatyKey = newFixation.GetDictionaryKey();
            if (!FixationsService.fixationSetToFixationListDictionary.ContainsKey(dictionatyKey))
                FixationsService.fixationSetToFixationListDictionary[dictionatyKey] = new List<Fixation>();
            FixationsService.fixationSetToFixationListDictionary[dictionatyKey].Add(newFixation);

            return newFixation;

        }

        public string GetDictionaryKey()
        {
            return this.Participant + '\t' + this.Trial + '\t' + this.Stimulus;
        }

        public string[] GetFixationDetailsAsArray()
        {
            string[] details = new string[12];

            details[0] = (this.Trial);
            details[1] = (this.Stimulus);
            details[2] = (this.Participant);
            details[3] = ("" + this.Index);
            details[4] = ("" + this.Event_Duration);
            details[5] = ("" + this.Fixation_Position_X);
            details[6] = ("" + this.Fixation_Position_Y);
            details[7] = ("" + this.Fixation_Average_Pupil_Diameter);
            details[8] = ("" + this.AOI_Group_Before_Change);
            details[9] = ("" + this.AOI_Group_After_Change);
            details[10] = ("" + this.AOI_Name);
            details[11] = ("" + (this.IsException ? 1 : 0));

            return details;
        }

        public bool ShouldBeSkippedInFirstPass()
        {
            bool isException = IsException;
            bool isInBoundAndShouldBeSkipped = IsInExceptionBounds && FixationsService.dealingWithInsideExceptions == DealingWithExceptionsEnum.Skip_In_First_Pass;
            bool isOutOfBoundAndShouldBeSkipped = !IsInExceptionBounds && FixationsService.dealingWithOutsideExceptions == DealingWithExceptionsOutBoundsEnum.Skip_In_First_Pass;
            return isException && (isInBoundAndShouldBeSkipped || isOutOfBoundAndShouldBeSkipped);
        }

        public double DistanceTo(Fixation other)
        {
            return Math.Sqrt(Math.Pow(this.Fixation_Position_X - other.Fixation_Position_X, 2) + Math.Pow(this.Fixation_Position_Y - other.Fixation_Position_Y, 2));
        }

        public double DistanceToPreviousFixation()
        {
            return this.DistanceTo(this.Previous_Fixation);
        }

        public bool IsBeforeThan(Fixation other)
        {
            return this.AOI_Name < other.AOI_Name || (this.Fixation_Position_X > other.Fixation_Position_X && this.AOI_Name == other.AOI_Name);
        }

        private static Log CreateLogForFieldValidation(string fieldName, string valueFound, uint lineNumber)
        {
            return new Log() { FileName = FixationsService.textFileName, LineNumber = lineNumber, Description = "The Value Of Field " + fieldName + " Is Not Valid!!! " + Environment.NewLine + "The Value Found Is: " + valueFound };
        }
    }
}
