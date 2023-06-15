using System.Runtime.Serialization;

namespace OutlookListener.AppointmentDetails
{
    [DataContract]
    public enum DayOfWeek
    {
        [EnumMember]
        Sunday = 1,
        [EnumMember]
        Monday = 2,
        [EnumMember]
        Tuesday = 4,
        [EnumMember]
        Wednesday = 8,
        [EnumMember]
        Thursday = 16,
        [EnumMember]
        Friday = 32,
        [EnumMember]
        Saturday = 64
    }
}
