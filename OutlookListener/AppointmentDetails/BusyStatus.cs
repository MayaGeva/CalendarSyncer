using System.Runtime.Serialization;

namespace OutlookListener.AppointmentDetails
{
    [DataContract]
    public enum BusyStatus
    {
        [EnumMember]
        Free,
        [EnumMember]
        Tentative,
        [EnumMember]
        Busy,
        [EnumMember]
        OutOfOffice,
        [EnumMember]
        WorkingElsewhere
    }
}
