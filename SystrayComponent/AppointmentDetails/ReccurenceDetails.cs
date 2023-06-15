using Microsoft.Office.Interop.Outlook;
using System.Runtime.Serialization;

namespace SystrayComponent.AppointmentDetails
{
    [DataContract]
    public class ReccurenceDetails
    {
        [DataMember]
        public int Day { get; set; }
        [DataMember]
        public int Month { get; set; }
        [DataMember]
        public ReccurenceType ReccurenceType { get; set; }
        [DataMember]
        public DayOfWeek DayOfWeek { get; set; }
        [DataMember]
        public int Interval { get; set; }
        [DataMember]
        public DateTime Start { get; set; }
        [DataMember]
        public DateTime End { get; set; }
        [DataMember]
        public bool NoEndDate { get; set; }
        [DataMember]
        public int Occurrences { get; set; }

        public ReccurenceDetails(RecurrencePattern reccurencePattern)
        {
            Day = reccurencePattern.DayOfMonth;
            Month = reccurencePattern.MonthOfYear;
            ReccurenceType = (ReccurenceType)reccurencePattern.RecurrenceType;
            Start = reccurencePattern.PatternStartDate;
            End = reccurencePattern.PatternEndDate;
            DayOfWeek = (DayOfWeek)reccurencePattern.DayOfWeekMask;
            Interval = reccurencePattern.Interval;
            Occurrences = reccurencePattern.Occurrences;
        }
    }
}
