using System.Runtime.Serialization;

namespace CalendarSyncer.AppointmentDetails
{
    public enum BusyStatus
    {
        [EnumMember]
        Busy,
        [EnumMember]
        Tentative,
        [EnumMember]
        Free,
        [EnumMember]
        OutOfOffice,
        [EnumMember]
        WorkingElsewhere
    }
}
