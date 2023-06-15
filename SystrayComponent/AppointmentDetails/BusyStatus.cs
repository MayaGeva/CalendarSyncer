using System.Runtime.Serialization;

namespace SystrayComponent.AppointmentDetails
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
