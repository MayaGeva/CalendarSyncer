using System.Reflection;
using Windows.ApplicationModel.Appointments;
using Windows.Globalization;
using Windows.Storage;

namespace SyncerApp.Calendar.Windows
{
    internal class WindowsCalendar
    {
        AppointmentStore? appointmentStore;
        AppointmentCalendar? appCalendar;
        const string OUTLOOK_CALENDAR = "OutlookCalendar";


        public WindowsCalendar()
        {
            InitCalendar();
        }

        async void InitCalendar()
        {
            appointmentStore = await AppointmentManager.RequestStoreAsync(AppointmentStoreAccessType.AllCalendarsReadWrite);
            string? calendarId = await GetCalendarIdFromName(OUTLOOK_CALENDAR);

            if (calendarId == null) {
                CreateCalendar();
            } else {
                appCalendar = await appointmentStore.GetAppointmentCalendarAsync(calendarId);
            }
        }

        async void CreateCalendar()
        {
            appCalendar = await appointmentStore?.CreateAppointmentCalendarAsync(OUTLOOK_CALENDAR);
            ApplicationData.Current.LocalSettings.Values[OUTLOOK_CALENDAR] = appCalendar.LocalId;
        }

        public async Task AddAppointment(Appointment appointment)
        {
            try {
                string? localId = await GetLocalIdFromRoamingId(appointment.RoamingId);
                await appCalendar?.SaveAppointmentAsync(appointment);
            } catch (ArgumentOutOfRangeException) {
                // Appointment with this roaming id exists in calendar
                await ModifyAppointment(appointment);
            }
        }

        public async Task RemoveAppointment(Appointment appointment)
        {
            string? localId = await GetLocalIdFromRoamingId(appointment.RoamingId);
            await appCalendar?.DeleteAppointmentAsync(localId);
        }

        public async Task ModifyAppointment(Appointment appointment)
        {
            string? localId = await GetLocalIdFromRoamingId(appointment.RoamingId);
            if (!string.IsNullOrEmpty(localId)) {
                Appointment original = await appointmentStore?.GetAppointmentAsync(localId);

                foreach (PropertyInfo property in typeof(Appointment).GetProperties().Where(p => p.CanWrite)) {
                    property.SetValue(original, property.GetValue(appointment, null), null);
                }
                await appCalendar?.SaveAppointmentAsync(original);
            } else {
                await AddAppointment(appointment);
            }
        }

        async Task<string?> GetLocalIdFromRoamingId(string roamingId)
        {
            IReadOnlyList<string> localIds = await appointmentStore?.FindLocalIdsFromRoamingIdAsync(roamingId);
            return localIds?.FirstOrDefault();
        }

        /// <summary>
        /// Get LocalId from calendar name
        /// </summary>
        /// <param name="calendarName">The name of the calendar</param>
        /// <returns>The LocalId if the calendar exists, or null otherwise</returns>
        async Task<string?> GetCalendarIdFromName(string calendarName)
        {
            var calendars = await appointmentStore?.FindAppointmentCalendarsAsync(FindAppointmentCalendarsOptions.IncludeHidden);
            List<string> calendarNames = calendars.Select(o => o.DisplayName).ToList();
            if (calendarNames.Contains(calendarName)) {
                return calendars[calendarNames.IndexOf(calendarName)].LocalId;
            }
            return null;
        }
    }
}
