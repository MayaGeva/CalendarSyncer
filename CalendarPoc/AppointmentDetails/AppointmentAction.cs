using System.Runtime.Serialization;

namespace CalendarSyncer.AppointmentDetails
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
