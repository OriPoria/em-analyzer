using ClosedXML.Attributes;
using ForBarIlanResearch.Enums;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ForBarIlanResearch.ModelClasses
{
    class Fixation
    {

        [XLColumn(Header = "Trial")]
        public string Trial { get; set; }
        [XLColumn(Header = "Stimulus")]
        public string Stimulus { get; set; }
        [XLColumn(Ignore = true)]
        public Fixation Previous_Fixation { get; set; }
        [XLColumn(Header = "Participant")]
        public string Participant { get; set; }
        [XLColumn(Header = "AOI Name")]
        public int AOI_Name { get; set; }
        [XLColumn(Header = "AOI Group Before Change")]
        public int AOI_Group_Before_Change { get; set; }
        [XLColumn(Header = "AOI Group After Change")]
        public int AOI_Group_After_Change { get; set; }
        [XLColumn(Ignore = true)]
        public long AOI_Size { get; set; }
        [XLColumn(Header = "AOI Coverage In Percents")]
        public double AOI_Coverage_In_Percents { get; set; }
        [XLColumn(Header = "Index")]
        public long Index { get; set; }
        [XLColumn(Header = "Event Duration")]
        public double Event_Duration { get; set; }
        [XLColumn(Header = "Fixation Position X")]
        public double Fixation_Position_X { get; set; }
        [XLColumn(Header = "Fixation Position Y")]
        public double Fixation_Position_Y { get; set; }
        [XLColumn(Header = "Fixation Average Pupil Diameter")]
        public double Fixation_Average_Pupil_Diameter { get; set; }
        [XLColumn(Header = "Is Exception")]
        public bool IsException { get; set; }
        [XLColumn(Ignore = true)]
        public AOIDetails AOI_Details { get; private set; }

        [XLColumn(Ignore = true)]
        public bool IsInExceptionBounds { get; set; }

        static double minimumDuration = double.Parse(ConfigurationService.getValue(ConfigurationService.Minimum_Event_Duration_In_ms));

        public static Fixation CreateFixationFromStringArray(string[] arr)
        {

            Fixation newFixation = new Fixation();
            newFixation.Event_Duration = double.Parse(arr[(int)TableColumnsEnum.Event_Duration]);

            if (newFixation.Event_Duration < minimumDuration)
                return null;

            try
            {
                newFixation.AOI_Group_After_Change = newFixation.AOI_Group_Before_Change = int.Parse(arr[(int)TableColumnsEnum.AOI_Group]);
                newFixation.AOI_Name = int.Parse(arr[(int)TableColumnsEnum.AOI_Name]);
                newFixation.AOI_Size = long.Parse(arr[(int)TableColumnsEnum.AOI_Size]);
            }
            catch
            {
                newFixation.AOI_Group_After_Change = newFixation.AOI_Group_Before_Change = -1;
                newFixation.AOI_Name = -1;
                newFixation.AOI_Size = -1;
            }

            try
            {
                newFixation.Fixation_Position_X = double.Parse(arr[(int)TableColumnsEnum.Fixation_Position_X]);
                newFixation.Fixation_Position_Y = double.Parse(arr[(int)TableColumnsEnum.Fixation_Position_Y]);
                newFixation.Fixation_Average_Pupil_Diameter = double.Parse(arr[(int)TableColumnsEnum.Fixation_Average_Pupil_Diameter]);
            }
            catch
            {
                return null;
            }

            newFixation.Trial = arr[(int)TableColumnsEnum.Trial].Trim();
            newFixation.Stimulus = arr[(int)TableColumnsEnum.Stimulus].Trim();
            newFixation.Participant = arr[(int)TableColumnsEnum.Participant].Trim();

            string dictionatyKey = newFixation.GetDictionaryKey();
            if (!FixationsService.fixationSetToFixationListDictionary.ContainsKey(dictionatyKey))
                FixationsService.fixationSetToFixationListDictionary[dictionatyKey] = new List<Fixation>();

            newFixation.AOI_Coverage_In_Percents = double.Parse(arr[(int)TableColumnsEnum.AOI_Coverage]);
            newFixation.Index = long.Parse(arr[(int)TableColumnsEnum.Index]);
            newFixation.IsException = false;

            if (newFixation.AOI_Name != -1)
            {
                if (AOIDetails.isAOIIncludeStimulus)
                    newFixation.AOI_Details = AOIDetails.nameToAOIDetailsDictionary[newFixation.AOI_Name + newFixation.Stimulus];
                else
                    newFixation.AOI_Details = AOIDetails.nameToAOIDetailsDictionary[newFixation.AOI_Name + ""];
            }

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

        public double DistanceTo(Fixation other)
        {
            return Math.Sqrt(Math.Pow(this.Fixation_Position_X - other.Fixation_Position_X, 2) + Math.Pow(this.Fixation_Position_Y - other.Fixation_Position_Y, 2));
        }

        public double DistanceToPreviousFixation()
        {
            return this.DistanceTo(this.Previous_Fixation);
        }

        public bool isBeforeThan(Fixation other)
        {
            return this.AOI_Name < other.AOI_Name || (this.Fixation_Position_X > other.Fixation_Position_X && this.AOI_Name == other.AOI_Name);
        }

        public double DistanceToAOI(AOIDetails aoi_details)
        {
            double right_border, left_border, upper_border, buttom_border;
            right_border = aoi_details.X + aoi_details.L / 2;
            left_border = aoi_details.X - aoi_details.L / 2;
            buttom_border = aoi_details.Y + aoi_details.H / 2;
            upper_border = aoi_details.Y - aoi_details.H / 2;

            if (this.Fixation_Position_X >= left_border && this.Fixation_Position_X <= right_border) //Inside X Bounds
            {
                if (this.Fixation_Position_Y >= upper_border && this.Fixation_Position_Y <= buttom_border)//Inside Y Bounds
                    return 0;
                else //Outside Y Bounds
                {
                    if (this.Fixation_Position_Y < upper_border)
                    {
                        return upper_border - this.Fixation_Position_Y;
                    }
                    else //this.Fixation_Position_Y > buttom_border
                    {
                        return this.Fixation_Position_Y - buttom_border;
                    }
                }
            }
            else //Outside X Bounds
            {
                if (this.Fixation_Position_Y >= upper_border && this.Fixation_Position_Y <= buttom_border)//Inside Y Bounds
                {
                    if (this.Fixation_Position_X < left_border)
                    {
                        return left_border - this.Fixation_Position_X;
                    }
                    else //this.Fixation_Position_X > right_border
                    {
                        return this.Fixation_Position_X - right_border;
                    }
                }
                else //Outside Y Bounds (And X Bounds)
                {
                    double Y_Distance, X_Distance;
                    if (this.Fixation_Position_Y < upper_border)
                    {
                        Y_Distance = upper_border - this.Fixation_Position_Y;
                    }
                    else //this.Fixation_Position_Y > buttom_border
                    {
                        Y_Distance = this.Fixation_Position_Y - buttom_border;
                    }

                    if (this.Fixation_Position_X < left_border)
                    {
                        X_Distance = left_border - this.Fixation_Position_X;
                    }
                    else //this.Fixation_Position_X > right_border
                    {
                        X_Distance = this.Fixation_Position_X - right_border;
                    }

                    return Math.Sqrt(Math.Pow(X_Distance, 2) + Math.Pow(Y_Distance, 2));
                }
            }
        }
    }
}
