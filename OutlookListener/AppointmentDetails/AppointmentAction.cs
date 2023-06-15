using System.Runtime.Serialization;

namespace OutlookListener.AppointmentDetails
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
