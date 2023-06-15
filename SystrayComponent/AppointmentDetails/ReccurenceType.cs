using System.Runtime.Serialization;

namespace SystrayComponent.AppointmentDetails
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
