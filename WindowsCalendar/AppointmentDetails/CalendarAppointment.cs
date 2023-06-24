using System;
using System.Runtime.Serialization;
using Windows.ApplicationModel.Appointments;

namespace WindowsCalendar.AppointmentDetails
{
    [DataContract]
    public class CalendarAppointment
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
            if (this.IsRecurring)
            {
                appointment.Recurrence = this.ReccurenceDetails.ToAppointmentRecurrence();
            }
            return appointment;
        }

        public CalendarAppointment(Appointment appointment)
        {
            this.Title = appointment.Subject;
            this.Description = appointment.Details;
            this.Location = appointment.Location;
            this.BusyStatus = (BusyStatus)appointment.BusyStatus;
            this.AllDayEvent = appointment.AllDay;
            this.Start = appointment.StartTime.DateTime;
            this.Duration = appointment.Duration;
            this.Organizer = appointment.Organizer.DisplayName;
            this.AppointmentId = appointment.RoamingId;
        }
        /*public bool Compare(Appointment other)
        {

        }*/
    }
}
