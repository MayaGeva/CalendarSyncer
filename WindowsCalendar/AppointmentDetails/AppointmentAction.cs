using System.Runtime.Serialization;

namespace WindowsCalendar.AppointmentDetails
{
    [DataContract]
    public enum AppointmentAction
    {
        [EnumMember]
        NoAction,
        [EnumMember]
        AddItem,
        [EnumMember]
        ChangeItem,
        [EnumMember]
        RemoveItem
    }
}
