using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Windows.ApplicationModel.Background;
using WindowsCalendar.AppointmentDetails;
using WindowsCalendar.Calendar;
using WindowsCalendar;

namespace WindowsCalendar
{
    internal class CalendarSyncTask : IBackgroundTask
    {
        BackgroundTaskDeferral _deferral;
        public void Run(IBackgroundTaskInstance taskInstance)
        {
            _deferral = taskInstance.GetDeferral();
            BlockingCollection<CalendarAppointment> calendarAppointments = new BlockingCollection<CalendarAppointment>();
            Task pipeServer = Task.Run(() => new PipeServer(calendarAppointments).Run());
            Task appointmentHandler = Task.Run(() => new WindowsAppCalendar(calendarAppointments).HandleAppointments());
            _deferral.Complete();
        }
    }
}
