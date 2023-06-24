using System;
using System.Runtime.Serialization;
using Windows.ApplicationModel.Appointments;

namespace CalendarTasks.AppointmentDetails
{
    [DataContract]
    public struct CalendarAppointment
    {
        [DataMember]
        public AppointmentAction Action;
        [DataMember]
        public string AppointmentId;
        [DataMember]
        public BusyStatus BusyStatus;
        [DataMember]
        public string Title;
        [DataMember]
        public string Description;
        [DataMember]
        public bool AllDayEvent;
        [DataMember]
        public TimeSpan Duration;
        [DataMember]
        public DateTime Start;
        [DataMember]
        public bool IsRecurring;
        [DataMember]
        public ReccurenceDetails ReccurenceDetails;
        [DataMember]
        public string Location;
        [DataMember]
        public string Organizer;

        public Appointment ToAppointment()
        {
            Appointment appointment = new Appointment()
            {
                Subject = this.Title,
                Details = this.Description,
                BusyStatus = (AppointmentBusyStatus)(int)this.BusyStatus,
                AllDay = this.AllDayEvent,
                StartTime = this.Start,
                Duration = this.Duration,
                Location = this.Location,
                Organizer = new AppointmentOrganizer() { DisplayName = this.Organizer },
                RoamingId = this.AppointmentId
            };
            if (this.IsRecurring )
            {
                appointment.Recurrence = this.ReccurenceDetails.ToAppointmentRecurrence();
            }
            return appointment;
        }
    }
}
