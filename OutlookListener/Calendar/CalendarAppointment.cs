using Microsoft.Office.Interop.Outlook;
using System.Runtime.Serialization;
using OutlookListener.AppointmentDetails;

namespace OutlookListener.Calendar
{
    [DataContract]
    internal class CalendarAppointment
    {

        [DataMember]
        public AppointmentAction Action { get; set; }
        [DataMember]
        public string AppointmentId { get; set; }
        [DataMember]
        public BusyStatus BusyStatus { get; set; }
        [DataMember]
        public string Title { get; set; }
        [DataMember]
        public string Description { get; set; }
        [DataMember]
        public bool AllDayEvent { get; set; }
        [DataMember]
        public TimeSpan Duration { get; set; }
        [DataMember]
        public DateTime Start { get; set; }
        [DataMember]
        public bool IsRecurring { get; set; }
        [DataMember]
        public ReccurenceDetails ReccurenceDetails { get; set; }
        [DataMember]
        public string Location { get; set; }
        [DataMember]
        public string Organizer { get; set; }
        public CalendarAppointment(AppointmentItem appointmentItem, AppointmentAction action = AppointmentAction.NoAction)
        {
            Action = action;
            AppointmentId = appointmentItem.EntryID;
            BusyStatus = (BusyStatus)(int)appointmentItem.BusyStatus;
            Title = appointmentItem.Subject;
            Description = appointmentItem.Body;
            AllDayEvent = appointmentItem.AllDayEvent;
            Duration = TimeSpan.FromMinutes(appointmentItem.Duration);
            Start = appointmentItem.Start;
            IsRecurring = appointmentItem.IsRecurring;
            ReccurenceDetails = new ReccurenceDetails(appointmentItem.GetRecurrencePattern());
            Location = appointmentItem.Location;
            Organizer = appointmentItem.Organizer;
        }
    }
}
