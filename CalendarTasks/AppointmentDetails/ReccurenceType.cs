using System.Runtime.Serialization;

namespace CalendarTasks.AppointmentDetails
{
    [DataContract]
    public enum ReccurenceType
    {
        [EnumMember]
        Daily,
        [EnumMember]
        Weekly,
        [EnumMember]
        Monthly,
        [EnumMember]
        MonthlyOnDay,
        [EnumMember]
        Yearly,
        [EnumMember]
        YearlyOnDay
    }
}
