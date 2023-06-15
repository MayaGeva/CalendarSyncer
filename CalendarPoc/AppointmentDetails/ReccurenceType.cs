﻿using System.Runtime.Serialization;

namespace CalendarSyncer.AppointmentDetails
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
        MonthlyOnDay,
        [EnumMember]
        Yearly,
        [EnumMember]
        YearlyOnDay
    }
}
