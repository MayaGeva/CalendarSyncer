using System.Collections.Concurrent;
using Windows.ApplicationModel.Appointments;
using Windows.Storage;

namespace SyncerApp.Calendar.Windows
{
    internal class WindowsCalendarSyncer
    {
        BlockingCollection<CalendarAppointment> appointmentCollection;
        AppointmentConverter appointmentConverter;
        WindowsCalendar windowsCalendar;
        const string LAST_SYNC_TIME = "LastSyncTime";

        public WindowsCalendarSyncer(
            BlockingCollection<CalendarAppointment> calendarAppointments, 
            AppointmentConverter appointmentConverter,
            WindowsCalendar windowsCalendar
            )
        {
            appointmentCollection = calendarAppointments;
            this.appointmentConverter = appointmentConverter;
            this.windowsCalendar = windowsCalendar;
        }

        public async Task HandleAppointments(CancellationToken token)
        {
            bool keep_running = true;
            while (keep_running)
            {
                if (token.IsCancellationRequested)
                {
                    keep_running = false;
                }
                CalendarAppointment calendarAppointment = await GetAppointment();
                Appointment appointment = new();
                if (calendarAppointment.Appointment != null)
                {
                    appointment = appointmentConverter.ToAppointment(calendarAppointment.Appointment);
                }
                switch (calendarAppointment.Action)
                {
                    case AppointmentAction.AddItem:
                        await windowsCalendar.AddAppointment(appointment);
                        break;
                    case AppointmentAction.RemoveItem:
                        await windowsCalendar.RemoveAppointment(appointment);
                        break;
                    case AppointmentAction.ChangeItem:
                        await windowsCalendar.ModifyAppointment(appointment);
                        break;
                    default:
                        ApplicationData.Current.LocalSettings.Values[LAST_SYNC_TIME] = DateTime.Now.ToString();
                        break;
                }
            }
        }
        Task<CalendarAppointment> GetAppointment()
        {
            CalendarAppointment calendarAppointment;
            Mutex mutex = new(true, "AppointmentsQueue", out bool createdNew);
            if (createdNew)
            {
                calendarAppointment = appointmentCollection.Take();
            }
            else
            {
                mutex.WaitOne();
                calendarAppointment = appointmentCollection.Take();
            }
            mutex.ReleaseMutex();
            return Task.FromResult(calendarAppointment);
        }
    }
}
