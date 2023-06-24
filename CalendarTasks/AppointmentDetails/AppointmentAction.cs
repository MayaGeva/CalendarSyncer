using System.Runtime.Serialization;

namespace CalendarTasks.AppointmentDetails
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
