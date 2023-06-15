using System;
using System.Runtime.Serialization;
using Windows.ApplicationModel.Appointments;

namespace CalendarSyncer.AppointmentDetails
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

        public AppointmentRecurrence ToAppointmentRecurrence()
        {
            AppointmentRecurrence recurrence = new AppointmentRecurrence()
            {
                Day = (uint)this.Day,
                Month = (uint)this.Month,
                DaysOfWeek = (AppointmentDaysOfWeek)this.DayOfWeek,
                Unit = (AppointmentRecurrenceUnit)this.ReccurenceType,
                Interval = (uint)this.Interval,
                Occurrences = (uint)this.Occurrences
            };
            return recurrence;
        }
    }
}
