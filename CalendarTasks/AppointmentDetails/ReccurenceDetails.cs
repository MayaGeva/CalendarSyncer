using System;
using System.Runtime.Serialization;
using Windows.ApplicationModel.Appointments;

namespace CalendarTasks.AppointmentDetails
{
    [DataContract]
    public struct ReccurenceDetails
    {
        [DataMember]
        public int Day;
        [DataMember]
        public int Month;
        [DataMember]
        public ReccurenceType ReccurenceType;
        [DataMember]
        public DayOfWeek DayOfWeek;
        [DataMember]
        public int Interval;
        [DataMember]
        public DateTime Start;
        [DataMember]
        public DateTime End;
        [DataMember]
        public bool NoEndDate;
        [DataMember]
        public int Occurrences;

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
