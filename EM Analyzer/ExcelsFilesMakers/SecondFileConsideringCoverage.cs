using System;
using System.Collections.Generic;
using System.Linq;
using static EM_Analyzer.ExcelsFilesMakers.SecondFileAfterProccessing;
using EM_Analyzer.ModelClasses;
using EM_Analyzer.Services;
using System.ComponentModel;

namespace EM_Analyzer.ExcelsFilesMakers
{
    public class SecondFileConsideringCoverage
    {
        private delegate double NumericExpression(AIOClassAfterCoverage value);
        //private delegate void SettingValue(AIOClassAfterCoverageForExcel AOI, double value);
        private delegate void SettingValue(AIOClassAfterCoverageForExcel AOI, string value);

        private static List<NumericExpression> GetNumericExpressions()
        {
            List<NumericExpression> filteringsExpressions = new List<NumericExpression>
            {
                aoi => aoi.Total_Fixation_Duration,
                aoi => aoi.Total_Fixation_Number,
                aoi => aoi.First_Fixation_Duration,
                aoi => aoi.First_Pass_Duration,
                aoi => aoi.First_Pass_Number,
                aoi => aoi.First_Pass_Progressive_Duration,
                aoi => aoi.First_Pass_Progressive_Number,
                aoi => aoi.First_Pass_Progressive_Duration_Overall,
                aoi => aoi.First_Pass_Progressive_Number_Overall,
                aoi => aoi.Total_First_Pass_Progressive_Duration,
                aoi => aoi.Total_First_Pass_Progressive_Number,
                aoi => aoi.Total_First_Pass_Progressive_Duration_Overall,
                aoi => aoi.Total_First_Pass_Progressive_Number_Overall,
                aoi => aoi.Total_First_Pass_Regressive_Duration,
                aoi => aoi.Total_First_Pass_Regressive_Number,
                aoi => aoi.Regression_Number,
                aoi => aoi.Regression_Duration,
                aoi => aoi.First_Regression_Duration,
                aoi => aoi.Pupil_Diameter,
                //aoi=>aoi.Total_Pupil_Diameter,
                //aoi=>aoi.Mean_AOI_Size,
                //aoi=>aoi.Total_AOI_Size

            };
            return filteringsExpressions;
        }

        private static List<SettingValue> GetSetters()
        {
            List<SettingValue> settingExpressions = new List<SettingValue>
            {
                (aoi, value) => aoi.Total_Fixation_Duration = value,
                (aoi, value) => aoi.Total_Fixation_Number = value,
                (aoi, value) => aoi.First_Fixation_Duration = value,
                (aoi, value) => aoi.First_Pass_Duration = value,
                (aoi, value) => aoi.First_Pass_Number = value,
                (aoi, value) => aoi.First_Pass_Progressive_Duration = value,
                (aoi, value) => aoi.First_Pass_Progressive_Number = value,
                (aoi, value) => aoi.First_Pass_Progressive_Duration_Overall = value,
                (aoi, value) => aoi.First_Pass_Progressive_Number_Overall = value,
                (aoi, value) => aoi.Total_First_Pass_Progressive_Duration = value,
                (aoi, value) => aoi.Total_First_Pass_Progressive_Number = value,
                (aoi, value) => aoi.Total_First_Pass_Progressive_Duration_Overall = value,
                (aoi, value) => aoi.Total_First_Pass_Progressive_Number_Overall = value,
                (aoi, value) => aoi.Total_First_Pass_Regressive_Duration = value,
                (aoi, value) => aoi.Total_First_Pass_Regressive_Number = value,
                (aoi, value) => aoi.Regression_Number = value,
                (aoi, value) => aoi.Regression_Duration = value,
                (aoi, value) => aoi.First_Regression_Duration = value,
                (aoi, value) => aoi.Pupil_Diameter = value,
                //(aoi, value) => aoi.Total_Pupil_Diameter = value,
                //(aoi, value) => aoi.Mean_AOI_Size = value,
                //(aoi, value) => aoi.Total_AOI_Size = value
            };
            return settingExpressions;
        }

        public static void MakeExcelFile()
        {
            double standardDevisionAllowed;
            try
            {
                standardDevisionAllowed = double.Parse(ConfigurationService.StandardDeviation);
            }
            catch
            {
                ExcelLogger.ExcelLoggerService.AddLog(new ExcelLogger.Log() { FileName = ConfigurationService.CONFIG_FILE, Description = "Standard Deviation Value Is Not A Number" });
                standardDevisionAllowed = 2.5;
            }

            List<NumericExpression> filteringsExpressions = GetNumericExpressions();
            List<SettingValue> settingExpressions = GetSetters();

            List<AIOClassAfterCoverageForExcel> byParticipant = DeleteOutOfStdValues(
                aoi => aoi.Participant,
                filteringsExpressions,
                settingExpressions,
                standardDevisionAllowed,
                aoi => aoi.AOI_Group == -1);
            ExcelsService.CreateExcelFromStringTable(ConfigurationService.ConsideredSecondExcelFileName + " By Participant", byParticipant);

            List<AIOClassAfterCoverageForExcel> byAOIGroup = DeleteOutOfStdValues(
                aoi => aoi.Stimulus + aoi.AOI_Group,
                filteringsExpressions,
                settingExpressions,
                standardDevisionAllowed,
                aoi => aoi.AOI_Group == -1);
            ExcelsService.CreateExcelFromStringTable(ConfigurationService.ConsideredSecondExcelFileName + " By AOI", byAOIGroup);
        }

        private static List<AIOClassAfterCoverageForExcel> DeleteOutOfStdValues(Func<AIOClassAfterCoverage, string> KeySelector,
            List<NumericExpression> filteringsExpressions,
            List<SettingValue> settingExpressions,
            double standardDevisionAllowed,
            Predicate<AIOClassAfterCoverage> AOIsToIgnore = null)
        {
            List<AIOClassAfterCoverage> acceptedAOIs = new List<AIOClassAfterCoverage>(AIOClassAfterCoverage.allInstances);
            if (AOIsToIgnore != null)
                acceptedAOIs.RemoveAll(AOIsToIgnore);

            var results = from aoi in acceptedAOIs
                          group aoi by KeySelector(aoi) into grou
                          select new { Key = grou.Key, AIOS = grou.ToList() };

            Dictionary<string, List<AIOClassAfterCoverage>> dict = new Dictionary<string, List<AIOClassAfterCoverage>>();
            results.ToList().ForEach(tuple => dict[tuple.Key] = tuple.AIOS);

            for (int fieldIndex = 0 ; fieldIndex < filteringsExpressions.Count ; fieldIndex++)
            {
                NumericExpression currentFilter = filteringsExpressions[fieldIndex];
                SettingValue currentSetter = settingExpressions[fieldIndex];
                foreach (string key in dict.Keys)
                {
                    List<AIOClassAfterCoverage> currentAOIs = dict[key];
                    IEnumerable<double> valuesForStandardDevision = currentAOIs.Select(aoi => currentFilter(aoi));
                    IEnumerable<double> standardDevisionGrades = StandardDevision.ComputeStandardDevisionGrades(valuesForStandardDevision);
                    int length = standardDevisionGrades.Count();
                    for (int i = 0 ; i < length ; ++i)
                    {
                        if (Math.Abs(standardDevisionGrades.ElementAt(i)) > standardDevisionAllowed)
                            currentSetter(currentAOIs[i].AOIForExcel, "");
                        else
                            currentSetter(currentAOIs[i].AOIForExcel, currentFilter(currentAOIs[i]) + "");
                    }
                }
            }
            List<AIOClassAfterCoverageForExcel> forExcel = acceptedAOIs.Select(aoi => aoi.AOIForExcel).ToList();
            return forExcel;
        }


        public class AIOClassAfterCoverage
        {
            public static List<AIOClassAfterCoverage> allInstances = new List<AIOClassAfterCoverage>();
            private readonly AOIClass AOI;
            public AIOClassAfterCoverageForExcel AOIForExcel;

            [Description("Participant")]
            public string Participant { get => AOI.Participant; }
            [Description("Trial")]
            public string Trial { get => AOI.Trial; }
            [Description("Stimulus")]
            public string Stimulus { get => AOI.Stimulus; }
            [Description("Text Name")]
            public string Text_Name { get => AOI.Text_Name; }
            [Description("AOI Group")]
            public int AOI_Group { get => AOI.AOI_Group; }

            [Description("Total Fixation Duration")]
            public double Total_Fixation_Duration
            {
                get
                {
                    return AOI.Total_Fixation_Duration / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("Total Fixation Number")]
            public double Total_Fixation_Number
            {
                get
                {
                    return AOI.Total_Fixation_Number / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("First Fixation Duration")]
            public double First_Fixation_Duration
            {
                get
                {
                    return AOI.First_Fixation_Duration / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("First-Pass Duration")]
            public double First_Pass_Duration
            {
                get
                {
                    return AOI.First_Pass_Duration / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("First-Pass Number")]
            public double First_Pass_Number
            {
                get
                {
                    return AOI.First_Pass_Number / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("First-Pass Progressive Duration")]
            public double First_Pass_Progressive_Duration
            {
                get
                {
                    return AOI.First_Pass_Progressive_Duration / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("First-Pass Progressive Number")]
            public double First_Pass_Progressive_Number
            {
                get
                {
                    return AOI.First_Pass_Progressive_Number / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("First-Pass Progressive Duration Overall")]
            public double First_Pass_Progressive_Duration_Overall
            {
                get
                {
                    return AOI.First_Pass_Progressive_Duration_Overall / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("First-Pass Progressive Number Overall")]
            public double First_Pass_Progressive_Number_Overall
            {
                get
                {
                    return AOI.First_Pass_Progressive_Number_Overall / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("Total First-Pass Progressive Duration")]
            public double Total_First_Pass_Progressive_Duration
            {
                get
                {
                    return AOI.Total_First_Pass_Progressive_Duration / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("Total First-Pass Progressive Number")]
            public double Total_First_Pass_Progressive_Number
            {
                get
                {
                    return AOI.Total_First_Pass_Progressive_Number / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("Total First-Pass Progressive Duration Overall")]
            public double Total_First_Pass_Progressive_Duration_Overall
            {
                get
                {
                    return AOI.Total_First_Pass_Progressive_Duration_Overall / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("Total First-Pass Progressive Number Overall")]
            public double Total_First_Pass_Progressive_Number_Overall
            {
                get
                {
                    return AOI.Total_First_Pass_Progressive_Number_Overall / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("Total First-Pass Regressive Duration")]
            public double Total_First_Pass_Regressive_Duration
            {
                get
                {
                    return AOI.Total_First_Pass_Regressive_Duration / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("Total First-Pass Regressive Number")]
            public double Total_First_Pass_Regressive_Number
            {
                get
                {
                    return AOI.Total_First_Pass_Regressive_Number / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("Regression Number")]
            public double Regression_Number
            {
                get
                {
                    return AOI.Regression_Number / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("Regression Duration")]
            public double Regression_Duration
            {
                get
                {
                    return AOI.Regression_Duration / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("First Regression Duration")]
            public double First_Regression_Duration
            {
                get
                {
                    return AOI.First_Regression_Duration / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("Skip")]
            public bool Skip { get; set; }

            [Description("Pupil Diameter [mm]")]
            public double Pupil_Diameter
            {
                get
                {
                    return AOI.Pupil_Diameter / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("AOI Size X [mm]")]
            public double Mean_AOI_Size
            {
                get
                {
                    return AOI.Mean_AOI_Size / AOI.Mean_AOI_Coverage;
                }
            }

            [Description("AOI Coverage [%]")]
            public double Mean_AOI_Coverage
            {
                get
                {
                    return AOI.Mean_AOI_Coverage;
                }
            }

            public AIOClassAfterCoverage(AOIClass AOI)
            {
                this.AOI = AOI;
                allInstances.Add(this);
                this.AOIForExcel = new AIOClassAfterCoverageForExcel(this);
            }

        }

        public class AIOClassAfterCoverageForExcel
        {
            private AIOClassAfterCoverage AOIAfterCoverage;

            [Description("Participant")]
            public string Participant { get => AOIAfterCoverage.Participant; }
            [Description("Trial")]
            public string Trial { get => AOIAfterCoverage.Trial; }
            [Description("Stimulus")]
            public string Stimulus { get => AOIAfterCoverage.Stimulus; }
            [Description("Text Name")]
            public string Text_Name { get => AOIAfterCoverage.Text_Name; }

            [Description("Total Fixation Duration")]
            public string Total_Fixation_Duration { get; set; }

            [Description("Total Fixation Number")]
            public string Total_Fixation_Number { get; set; }

            [Description("First Fixation Duration")]
            public string First_Fixation_Duration { get; set; }

            [Description("First-Pass Duration")]
            public string First_Pass_Duration { get; set; }

            [Description("First-Pass Number")]
            public string First_Pass_Number { get; set; }

            [Description("First-Pass Progressive Duration")]
            public string First_Pass_Progressive_Duration { get; set; }

            [Description("First-Pass Progressive Number")]
            public string First_Pass_Progressive_Number { get; set; }

            [Description("First-Pass Progressive Duration Overall")]
            public string First_Pass_Progressive_Duration_Overall { get; set; }

            [Description("First-Pass Progressive Number Overall")]
            public string First_Pass_Progressive_Number_Overall { get; set; }

            [Description("Total First-Pass Progressive Duration")]
            public string Total_First_Pass_Progressive_Duration { get; set; }

            [Description("Total First-Pass Progressive Number")]
            public string Total_First_Pass_Progressive_Number { get; set; }

            [Description("Total First-Pass Progressive Duration Overall")]
            public string Total_First_Pass_Progressive_Duration_Overall { get; set; }

            [Description("Total First-Pass Progressive Number Overall")]
            public string Total_First_Pass_Progressive_Number_Overall { get; set; }

            [Description("Total First-Pass Regressive Duration")]
            public string Total_First_Pass_Regressive_Duration { get; set; }

            [Description("Total First-Pass Regressive Number")]
            public string Total_First_Pass_Regressive_Number { get; set; }

            [Description("Regression Number")]
            public string Regression_Number { get; set; }

            [Description("Regression Duration")]
            public string Regression_Duration { get; set; }

            [Description("First Regression Duration")]
            public string First_Regression_Duration { get; set; }

            [Description("Skip")]
            public bool Skip { get; set; }

            [Description("Pupil Diameter [mm]")]
            public string Pupil_Diameter { get; set; }

            [Description("AOI Size X [mm]")]
            public string Mean_AOI_Size { get; set; }

            public AIOClassAfterCoverageForExcel(AIOClassAfterCoverage AOIClassAfterCoverage)
            {
                this.AOIAfterCoverage = AOIClassAfterCoverage;
            }
        }
    }
}
