using System;
using System.Collections.Generic;
using System.Linq;
using static EM_Analyzer.ExcelsFilesMakers.SecondFileAfterProccessing;
using EM_Analyzer.ModelClasses;
using EM_Analyzer.Services;
using System.ComponentModel;
using EM_Analyzer.Interfaces;
using System.Globalization;
using OfficeOpenXml;
using EM_Analyzer.Enums;

namespace EM_Analyzer.ExcelsFilesMakers
{
    public class SecondFileConsideringCoverage
    {
        private delegate double NumericExpression(AIOClassAfterCoverage value);
        private delegate void SettingValue(AIOClassAfterCoverageForExcel AOI, string value);
        public static AOITypes currentType;

        private static List<NumericExpression> GetNumericExpressions()
        {
            // list of functions that get AIOClassAfterCoverage as parameter and return double
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
                aoi => aoi.Pupil_Diameter

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
            ExcelsService.CreateExcelFromStringTable(ConfigurationService.ConsideredSecondExcelFileName + " By Participant" + "_" + Constans.GetEndOfFileNameByType(currentType), byParticipant, EditExcel);

            List<AIOClassAfterCoverageForExcel> byAOIGroup = DeleteOutOfStdValues(
                aoi => aoi.Stimulus + aoi.AOI_Group,
                filteringsExpressions,
                settingExpressions,
                standardDevisionAllowed,
                aoi => aoi.AOI_Group == -1);
            ExcelsService.CreateExcelFromStringTable(ConfigurationService.ConsideredSecondExcelFileName + " By AOI" + "_" + Constans.GetEndOfFileNameByType(currentType), byAOIGroup, EditExcel);


            HashSet<AIOClassAfterCoverageForExcel> by_AIO_And_Participant = new HashSet<AIOClassAfterCoverageForExcel>();
            by_AIO_And_Participant.UnionWith(byParticipant);
            by_AIO_And_Participant.UnionWith(byAOIGroup);
            ExcelsService.CreateExcelFromStringTable(ConfigurationService.ConsideredSecondExcelFileName + " By AOI and Participant" + "_" + Constans.GetEndOfFileNameByType(currentType), by_AIO_And_Participant, EditExcel);

            
            List<AIOClassAfterCoverageForExcel> by_AIO_Or_Participant = new List<AIOClassAfterCoverageForExcel>(Math.Min(byParticipant.Count, byAOIGroup.Count));
            foreach(AIOClassAfterCoverageForExcel group in byAOIGroup)
            {
                foreach(AIOClassAfterCoverageForExcel participant in byParticipant)
                {
                    if(group == participant)
                    {
                        by_AIO_Or_Participant.Add(group);
                    }
                }
            }
            ExcelsService.CreateExcelFromStringTable(ConfigurationService.ConsideredSecondExcelFileName + " By AOI or Participant" + "_" + Constans.GetEndOfFileNameByType(currentType), by_AIO_Or_Participant, EditExcel);
           
        }
        public static int EditExcel(ExcelWorksheet ws)
        {
            ws.InsertColumn(Constans.secondFileStartCondsInx, AOIClass.maxConditions);
            for (int i = Constans.secondFileStartCondsInx; i < Constans.secondFileStartCondsInx + AOIClass.maxConditions; i++)
                ws.Cells[1, i].Value = $"Cond{i - Constans.secondFileStartCondsInx + 1}";
            for (int i = 2; i <= ws.Dimension.Rows; i++)
            {
                if (ws.Cells[i, Constans.secondFileAoiTargetCol].Value != null)
                {
                    List<string> sNames = Constans.parseSpecialName(ws.Cells[i, Constans.secondFileAoiTargetCol].Value.ToString());
                    int k = 0;
                    for (int j = Constans.secondFileStartCondsInx; j < sNames.Count + Constans.secondFileStartCondsInx; j++)
                    {
                        ws.Cells[i, j].Value = sNames[k];
                        k++;
                    }
                }
                // change figure AOI group label from 0 to "figure"
                if (ws.Cells[i, Constans.secondFileAoiGroupCol].Value.ToString() == "0")
                    ws.Cells[i, Constans.secondFileAoiGroupCol].Value = "figure";

            }
            // removes frequency and length column
            if (currentType == AOITypes.Phrases)
            {
                int numberCols = ws.Dimension.Columns;
                ws.DeleteColumn(numberCols);
                ws.DeleteColumn(numberCols - 1);
            }
            return 0;
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
                    double afterFilter;
                    string valueString = " ";
                    for (int i = 0 ; i < length ; ++i)
                    {
                        if (Math.Abs(standardDevisionGrades.ElementAt(i)) > standardDevisionAllowed)
                            currentSetter(currentAOIs[i].AOIForExcel, "");
                        else
                        {
                            afterFilter = currentFilter(currentAOIs[i]);
                            valueString = afterFilter.ToString("G17");

                            currentSetter(currentAOIs[i].AOIForExcel, valueString);
                        }
                    }
                }
            }
            
            List<AIOClassAfterCoverageForExcel> forExcel = acceptedAOIs.Select(aoi => aoi.AOIForExcel).ToList();
            return forExcel;
        }
        


        public class AIOClassAfterCoverage
        {
            public static List<AIOClassAfterCoverage> allInstances = new List<AIOClassAfterCoverage>();
            private readonly IAOIClassForConsideringCoverage AOI;
            public AIOClassAfterCoverageForExcel AOIForExcel;
            
            public static int denominator_value = int.Parse(ConfigurationService.SecondFileFilteringDenominator);

            [Description("Participant")]
            public string Participant { get => AOI.Participant; }

            [Description("Stimulus")]
            public string Stimulus { get => AOI.Stimulus; }
            [Description("Text Name")]
            public string Text_Name { get => AOI.Text_Name; }
            [Description("AOI Group")]
            public int AOI_Group { get => AOI.AOI_Group; }
            [Description("AOI Target")]
            public string AOI_Target { get => AOI.AOI_Target; }

            [Description("Total Fixation Duration")]
            public double Total_Fixation_Duration
            {
                get
                {
                    if(denominator_value == 2)
                        return AOI.Total_Fixation_Duration / AOI.Mean_AOI_Coverage;
                    if(denominator_value == 1)
                        return AOI.Total_Fixation_Duration / AOI.Mean_AOI_Size;
                    return AOI.Total_Fixation_Duration;
                }
            }

            [Description("Total Fixation Number")]
            public double Total_Fixation_Number
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.Total_Fixation_Number / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.Total_Fixation_Number / AOI.Mean_AOI_Size;
                    return AOI.Total_Fixation_Number;
                }
            }

            [Description("First Fixation Duration")]
            public double First_Fixation_Duration
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.First_Fixation_Duration / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.First_Fixation_Duration / AOI.Mean_AOI_Size;
                    return AOI.First_Fixation_Duration;
                }
            }

            [Description("First-Pass Duration")]
            public double First_Pass_Duration
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.First_Pass_Duration / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.First_Pass_Duration / AOI.Mean_AOI_Size;
                    return AOI.First_Pass_Duration;
                }
            }
            [Description("First-Pass Number")]
            public double First_Pass_Number {
                get
                {
                    if (denominator_value == 2)
                        return AOI.First_Pass_Number / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.First_Pass_Number / AOI.Mean_AOI_Size;
                    return AOI.First_Pass_Number;
                }
            }

            [Description("First-Pass Progressive Duration")]
            public double First_Pass_Progressive_Duration {
                get
                {
                    if (denominator_value == 2)
                        return AOI.First_Pass_Progressive_Duration / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.First_Pass_Progressive_Duration / AOI.Mean_AOI_Size;
                    return AOI.First_Pass_Progressive_Duration;
                }
            }

            [Description("First-Pass Progressive Number")]
            public double First_Pass_Progressive_Number
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.First_Pass_Progressive_Number / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.First_Pass_Progressive_Number / AOI.Mean_AOI_Size;
                    return AOI.First_Pass_Progressive_Number;
                }
            }

            [Description("First-Pass Progressive Duration Overall")]
            public double First_Pass_Progressive_Duration_Overall
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.First_Pass_Progressive_Duration_Overall / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.First_Pass_Progressive_Duration_Overall / AOI.Mean_AOI_Size;
                    return AOI.First_Pass_Progressive_Duration_Overall;
                }
            }

            [Description("First-Pass Progressive Number Overall")]
            public double First_Pass_Progressive_Number_Overall
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.First_Pass_Progressive_Number_Overall / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.First_Pass_Progressive_Number_Overall / AOI.Mean_AOI_Size;
                    return AOI.First_Pass_Progressive_Number_Overall;
                }
            }

            [Description("Total First-Pass Progressive Duration")]
            public double Total_First_Pass_Progressive_Duration
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.Total_First_Pass_Progressive_Duration / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.Total_First_Pass_Progressive_Duration / AOI.Mean_AOI_Size;
                    return AOI.Total_First_Pass_Progressive_Duration;
                }
            }

            [Description("Total First-Pass Progressive Number")]
            public double Total_First_Pass_Progressive_Number
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.Total_First_Pass_Progressive_Number / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.Total_First_Pass_Progressive_Number / AOI.Mean_AOI_Size;
                    return AOI.Total_First_Pass_Progressive_Number;
                }
            }

            [Description("Total First-Pass Progressive Duration Overall")]
            public double Total_First_Pass_Progressive_Duration_Overall
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.Total_First_Pass_Progressive_Duration_Overall / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.Total_First_Pass_Progressive_Duration_Overall / AOI.Mean_AOI_Size;
                    return AOI.Total_First_Pass_Progressive_Duration_Overall;
                }
            }

            [Description("Total First-Pass Progressive Number Overall")]
            public double Total_First_Pass_Progressive_Number_Overall
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.Total_First_Pass_Progressive_Number_Overall / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.Total_First_Pass_Progressive_Number_Overall / AOI.Mean_AOI_Size;
                    return AOI.Total_First_Pass_Progressive_Number_Overall;
                }
            }

            [Description("Total First-Pass Regressive Duration")]
            public double Total_First_Pass_Regressive_Duration
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.Total_First_Pass_Regressive_Duration / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.Total_First_Pass_Regressive_Duration / AOI.Mean_AOI_Size;
                    return AOI.Total_First_Pass_Regressive_Duration;
                }
            }

            [Description("Total First-Pass Regressive Number")]
            public double Total_First_Pass_Regressive_Number
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.Total_First_Pass_Regressive_Number / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.Total_First_Pass_Regressive_Number / AOI.Mean_AOI_Size;
                    return AOI.Total_First_Pass_Regressive_Number;
                }
            }

            [Description("Regression Number")]
            public double Regression_Number
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.Regression_Number / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.Regression_Number / AOI.Mean_AOI_Size;
                    return AOI.Regression_Number;
                }
            }

            [Description("Regression Duration")]
            public double Regression_Duration
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.Regression_Duration / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.Regression_Duration / AOI.Mean_AOI_Size;
                    return AOI.Regression_Duration;
                }
            }

            [Description("First Regression Duration")]
            public double First_Regression_Duration
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.First_Regression_Duration / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.First_Regression_Duration / AOI.Mean_AOI_Size;
                    return AOI.First_Regression_Duration;
                }
            }

            [Description("Skip")]
            public bool Skip
            { 
                get
                {
                    return AOI.Skip;
                }
            }

            [Description("Pupil Diameter [mm]")]
            public double Pupil_Diameter
            {
                get
                {
                    if (denominator_value == 2)
                        return AOI.Pupil_Diameter / AOI.Mean_AOI_Coverage;
                    if (denominator_value == 1)
                        return AOI.Pupil_Diameter / AOI.Mean_AOI_Size;
                    return AOI.Pupil_Diameter;
                }
            }

            public double Mean_AOI_Size
            {
                get
                {
                    return AOI.Mean_AOI_Size;
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
            public int Length
            {
                get
                {
                    return AOI.Length;
                }
            }
            public int Frequency
            {
                get
                {
                    return AOI.Frequency;
                }
            }
            public AIOClassAfterCoverage(IAOIClassForConsideringCoverage AOI)
            {
                this.AOI = AOI;
                allInstances.Add(this);
                this.AOIForExcel = new AIOClassAfterCoverageForExcel(this);
            }

        }

        public class AIOClassAfterCoverageForExcel
        {
            private AIOClassAfterCoverage AOIAfterCoverage;
            public static int maxConditions;


            [Description("Participant")]
            public string Participant { get => AOIAfterCoverage.Participant; }

            [Description("Stimulus")]
            public string Stimulus { get => AOIAfterCoverage.Stimulus; }
            [Description("Text Name")]
            public string Text_Name { get => AOIAfterCoverage.Text_Name; }

            [Description("AOI Group")]
            public int AOI_Group { get => AOIAfterCoverage.AOI_Group; }
            [Description("AOI Target")]
            public string AOI_Target { get => AOIAfterCoverage.AOI_Target; }

            [Description("Total Fixation Duration")]
            public string Total_Fixation_Duration { get;set; }

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
            public bool Skip { get => AOIAfterCoverage.Skip; }

            [Description("Pupil Diameter [mm]")]
            public string Pupil_Diameter { get; set; }
            [Description("Length")]
            public int Length { get => AOIAfterCoverage.Length; }
            [Description("Frequency")]
            public int Frequency { get => AOIAfterCoverage.Frequency; }
            public AIOClassAfterCoverageForExcel(AIOClassAfterCoverage AOIClassAfterCoverage)
            {
                this.AOIAfterCoverage = AOIClassAfterCoverage;
            }

            public static bool operator == (AIOClassAfterCoverageForExcel lhs, AIOClassAfterCoverageForExcel rhs)
            {
                if (lhs.Participant == rhs.Participant && lhs.Stimulus == rhs.Stimulus
                    && lhs.Text_Name == rhs.Text_Name && lhs.AOI_Group == rhs.AOI_Group && lhs.Total_Fixation_Duration == rhs.Total_Fixation_Duration &&
                    lhs.Total_Fixation_Number == rhs.Total_Fixation_Number && lhs.First_Fixation_Duration == rhs.First_Fixation_Duration && lhs.First_Pass_Duration == rhs.First_Pass_Duration
                        && lhs.First_Pass_Number == rhs.First_Pass_Number && lhs.First_Pass_Progressive_Duration == rhs.First_Pass_Progressive_Duration && lhs.First_Pass_Progressive_Number == rhs.First_Pass_Progressive_Number &&
                        lhs.First_Pass_Progressive_Duration_Overall == rhs.First_Pass_Progressive_Duration_Overall && lhs.First_Pass_Progressive_Number_Overall == rhs.First_Pass_Progressive_Number_Overall &&
                        lhs.Total_First_Pass_Progressive_Duration == rhs.Total_First_Pass_Progressive_Duration && lhs.Total_First_Pass_Progressive_Number == rhs.Total_First_Pass_Progressive_Number &&
                        lhs.Total_First_Pass_Progressive_Duration_Overall == rhs.Total_First_Pass_Progressive_Duration_Overall && lhs.Total_First_Pass_Progressive_Number_Overall == rhs.Total_First_Pass_Progressive_Number_Overall &&
                        lhs.Total_First_Pass_Regressive_Duration == rhs.Total_First_Pass_Regressive_Duration && lhs.Total_First_Pass_Regressive_Number == rhs.Total_First_Pass_Regressive_Number &&
                        lhs.Regression_Number == rhs.Regression_Number && lhs.Regression_Duration == rhs.Regression_Duration && lhs.First_Regression_Duration == rhs.First_Regression_Duration &&
                        lhs.Skip == rhs.Skip && lhs.Pupil_Diameter == rhs.Pupil_Diameter)
                    return true;
                return false;
            }
            public static bool operator != (AIOClassAfterCoverageForExcel lhs, AIOClassAfterCoverageForExcel rhs)
            {
                if (lhs.Participant == rhs.Participant  && lhs.Stimulus == rhs.Stimulus
                   && lhs.Text_Name == rhs.Text_Name && lhs.AOI_Group == rhs.AOI_Group && lhs.Total_Fixation_Duration == rhs.Total_Fixation_Duration &&
                   lhs.Total_Fixation_Number == rhs.Total_Fixation_Number && lhs.First_Fixation_Duration == rhs.First_Fixation_Duration && lhs.First_Pass_Duration == rhs.First_Pass_Duration
                       && lhs.First_Pass_Number == rhs.First_Pass_Number && lhs.First_Pass_Progressive_Duration == rhs.First_Pass_Progressive_Duration && lhs.First_Pass_Progressive_Number == rhs.First_Pass_Progressive_Number &&
                       lhs.First_Pass_Progressive_Duration_Overall == rhs.First_Pass_Progressive_Duration_Overall && lhs.First_Pass_Progressive_Number_Overall == rhs.First_Pass_Progressive_Number_Overall &&
                       lhs.Total_First_Pass_Progressive_Duration == rhs.Total_First_Pass_Progressive_Duration && lhs.Total_First_Pass_Progressive_Number == rhs.Total_First_Pass_Progressive_Number &&
                       lhs.Total_First_Pass_Progressive_Duration_Overall == rhs.Total_First_Pass_Progressive_Duration_Overall && lhs.Total_First_Pass_Progressive_Number_Overall == rhs.Total_First_Pass_Progressive_Number_Overall &&
                       lhs.Total_First_Pass_Regressive_Duration == rhs.Total_First_Pass_Regressive_Duration && lhs.Total_First_Pass_Regressive_Number == rhs.Total_First_Pass_Regressive_Number &&
                       lhs.Regression_Number == rhs.Regression_Number && lhs.Regression_Duration == rhs.Regression_Duration && lhs.First_Regression_Duration == rhs.First_Regression_Duration &&
                       lhs.Skip == rhs.Skip && lhs.Pupil_Diameter == rhs.Pupil_Diameter)
                    return false;
                return true;
            }
        }
    }
}
