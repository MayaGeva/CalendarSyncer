using System.Reflection;
using Windows.ApplicationModel.Appointments;
using Windows.Storage;

namespace SyncerApp.Calendar.Windows
{
    internal class WindowsCalendar
    {
        AppointmentStore? appointmentStore;
        AppointmentCalendar? appCalendar;
        const string OUTLOOK_CALENDAR = "OutlookCalendar";
        const string LAST_SYNC_TIME = "lastSyncTime";


        public WindowsCalendar()
        {
            InitCalendar();
        }

        async void InitCalendar()
        {
            bool firstTime = !ApplicationData.Current.LocalSettings.Values.ContainsKey(LAST_SYNC_TIME);
            appointmentStore = await AppointmentManager.RequestStoreAsync(AppointmentStoreAccessType.AllCalendarsReadWrite);
            
            string calendarId = (string)ApplicationData.Current.LocalSettings.Values[OUTLOOK_CALENDAR];
            appCalendar = await appointmentStore?.GetAppointmentCalendarAsync(calendarId);
            if (firstTime)
            {
                appCalendar?.DeleteAsync();
                CreateCalendar();
            }
        }

        async void CreateCalendar()
        {
            appCalendar = await appointmentStore?.CreateAppointmentCalendarAsync(OUTLOOK_CALENDAR);
            ApplicationData.Current.LocalSettings.Values[OUTLOOK_CALENDAR] = appCalendar.LocalId;
        }

        public async Task AddAppointment(Appointment appointment)
        {
            try
            {
                string? localId = await GetLocalIdFromRoamingId(appointment.RoamingId);
                await appCalendar?.SaveAppointmentAsync(appointment);
            }
            catch (ArgumentOutOfRangeException)
            {

            }
            catch (ArgumentNullException)
            {

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
            if (!string.IsNullOrEmpty(localId))
            {
                Appointment original = await appointmentStore?.GetAppointmentAsync(localId);
                foreach (PropertyInfo property in typeof(Appointment).GetProperties().Where(p => p.CanWrite))
                {
                    property.SetValue(original, property.GetValue(appointment, null), null);
                }
                await appCalendar?.SaveAppointmentAsync(original);
            }
            else
            {
                await AddAppointment(appointment);
            }
        }

        async Task<string?> GetLocalIdFromRoamingId(string roamingId)
        {
            IReadOnlyList<string> localIds = await appointmentStore?.FindLocalIdsFromRoamingIdAsync(roamingId);
            return localIds?.FirstOrDefault();
        }

        public async Task<List<Appointment>> GetEventsOnDate(DateTime date)
        {
            appointmentStore = await AppointmentManager.RequestStoreAsync(AppointmentStoreAccessType.AllCalendarsReadWrite);
            appCalendar = await appointmentStore?.CreateAppointmentCalendarAsync(OUTLOOK_CALENDAR);
            IReadOnlyList<Appointment> appointments = await appointmentStore?.FindAppointmentsAsync(new DateTimeOffset(date.Date), TimeSpan.FromHours(24));
            return new List<Appointment>(appointments);
        }
    }
}
