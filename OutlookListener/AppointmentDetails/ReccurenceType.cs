using System.Runtime.Serialization;

namespace OutlookListener.AppointmentDetails
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
        MonthlyNth,
        [EnumMember]
        Yearly,
        [EnumMember]
        YearlyNth
    }
}
