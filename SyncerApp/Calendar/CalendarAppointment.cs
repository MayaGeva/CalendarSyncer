using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
