using Microsoft.Office.Interop.Outlook;
using SyncerApp.Calendar.Converters;
using Windows.ApplicationModel.Appointments;

namespace SyncerApp.Calendar
{
    internal class AppointmentConverter
    {
        readonly Recipient currentUser;
        const string APPOINTMENT_URI = "ms-outlook://calendar/appointment?local_id={0}";
        public AppointmentConverter(Recipient currentUser) 
        {
            this.currentUser = currentUser;
        }
        public Appointment ToAppointment(AppointmentItem appointmentItem)
        {
            Appointment appointment = new()
            {
                Subject = appointmentItem.Subject,
                Details = appointmentItem.Body,
                DetailsKind = BodyFormatConverter.ConvertBodyFormat(appointmentItem.BodyFormat),
                AllDay = appointmentItem.AllDayEvent,
                StartTime = appointmentItem.Start,
                Duration = TimeSpan.FromMinutes(appointmentItem.Duration),
                BusyStatus = BusyStatusConverter.ConvertBusyStatus(appointmentItem.BusyStatus),
                Location = appointmentItem.Location,
                UserResponse = ResponseStatusConverter.ConvertResponseStatus(appointmentItem.ResponseStatus),
                RoamingId = appointmentItem.GlobalAppointmentID,
                Reminder = TimeSpan.FromMinutes(appointmentItem.ReminderMinutesBeforeStart),
                Uri = new(string.Format(APPOINTMENT_URI, appointmentItem.GlobalAppointmentID)),
            };
            if (appointmentItem.IsRecurring)
            {
                appointment.Recurrence = RecurrenceConverter.ConvertRecurrence(appointmentItem);
            }
            if (appointmentItem.Organizer == currentUser.Address)
            {
                appointment.IsOrganizedByUser = true;
            }
            else
            {
                appointment.Organizer = OrganizerConverter.ConvertOrganizer(appointmentItem.GetOrganizer());
            }
            return appointment;
        }   
    }
}
