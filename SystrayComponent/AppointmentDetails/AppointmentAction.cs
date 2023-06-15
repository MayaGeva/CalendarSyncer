using System.Runtime.Serialization;

namespace SystrayComponent.AppointmentDetails
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
