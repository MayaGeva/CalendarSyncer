using SyncerApp.Calendar;
using System.Collections.Concurrent;
using Windows.ApplicationModel.Appointments;
using Windows.Storage;

namespace SyncerApp.Calendar.Windows
{
    internal class WindowsCalendar
    {
        AppointmentStore appointmentStore;
        AppointmentCalendar appCalendar;

        CalendarStorageSettings calendarStorageSettings;
        const string OUTLOOK_CALENDAR = "OutlookCalendar";


        public WindowsCalendar(CalendarStorageSettings storageSettings)
        {
            calendarStorageSettings = storageSettings;
            InitCalendar();
        }

        async Task InitCalendar()
        {
            appointmentStore = await AppointmentManager.RequestStoreAsync(AppointmentStoreAccessType.AllCalendarsReadWrite);
            bool calendar_exists = ApplicationData.Current.LocalSettings.Values.ContainsKey(OUTLOOK_CALENDAR);
            if (calendar_exists)
            {
                string calendarId = (string)ApplicationData.Current.LocalSettings.Values[OUTLOOK_CALENDAR];
                appCalendar = await appointmentStore.GetAppointmentCalendarAsync(calendarId);
            }
            else
            {
                CreateCalendar();
            }
        }

        async void CreateCalendar()
        {
            appCalendar = await appointmentStore.CreateAppointmentCalendarAsync(OUTLOOK_CALENDAR);
            ApplicationData.Current.LocalSettings.Values[OUTLOOK_CALENDAR] = appCalendar.LocalId;
        }

        public async Task AddAppointment(Appointment appointment)
        {
            if (!calendarStorageSettings.RoamingIdExists(appointment.RoamingId))
            {
                await appCalendar.SaveAppointmentAsync(appointment);
                calendarStorageSettings.AddLocalIdMapping(appointment.RoamingId, appointment.LocalId);
            }
        }

        public async Task RemoveAppointment(Appointment appointment)
        {
            await appCalendar.DeleteAppointmentAsync(appointment.LocalId);
            calendarStorageSettings.RemoveLocalIdMapping(appointment.RoamingId);
        }

        public async Task ModifyAppointment(Appointment appointment)
        {
            string localId = calendarStorageSettings.GetLocalIdFromRoamingId(appointment.RoamingId);
            await appCalendar.DeleteAppointmentAsync(localId);
            await appCalendar.SaveAppointmentAsync(appointment);
            calendarStorageSettings.RemoveLocalIdMapping(localId, appointment.RoamingId);
            calendarStorageSettings.AddLocalIdMapping(appointment.RoamingId, appointment.LocalId);
        }

        public async Task<List<Appointment>> GetEventsOnDate(DateTime date)
        {
            appointmentStore = await AppointmentManager.RequestStoreAsync(AppointmentStoreAccessType.AllCalendarsReadWrite);
            appCalendar = await appointmentStore.CreateAppointmentCalendarAsync(OUTLOOK_CALENDAR);
            IReadOnlyList<Appointment> appointments = await appointmentStore.FindAppointmentsAsync(new DateTimeOffset(date.Date), TimeSpan.FromHours(24));
            return new List<Appointment>(appointments);
        }
    }
}
