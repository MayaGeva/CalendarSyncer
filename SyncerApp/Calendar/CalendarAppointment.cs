using Microsoft.Office.Interop.Outlook;

namespace SyncerApp.Calendar
{
    internal class CalendarAppointment
    {
        public AppointmentAction Action { get; set; }
        public AppointmentItem Appointment { get; set; }
        public CalendarAppointment(AppointmentAction action, AppointmentItem appointment)
        {
            Action = action;
            Appointment = appointment;
        }
    }
}
