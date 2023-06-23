using CalendarPoc;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading.Tasks;
using Windows.ApplicationModel.Appointments;
using Windows.Storage;
using WindowsCalendar.AppointmentDetails;

namespace WindowsCalendar.Calendar
{
    internal class WindowsAppCalendar
    {
        AppointmentStore appointmentStore;
        AppointmentCalendar appCalendar;

        BlockingCollection<CalendarAppointment> appointmentCollection;
        CalendarStorageSettings calendarStorageSettings;
        const string OUTLOOK_CALENDAR = "OutlookCalendar";
        
        public WindowsAppCalendar(BlockingCollection<CalendarAppointment> calendarAppointments) 
        {
            appointmentCollection = calendarAppointments;
            calendarStorageSettings = new CalendarStorageSettings();
            InitCalendar();
        }

        async Task InitCalendar()
        {
            this.appointmentStore = await AppointmentManager.RequestStoreAsync(AppointmentStoreAccessType.AllCalendarsReadWrite);
            bool calendar_exists = ApplicationData.Current.LocalSettings.Values.ContainsKey(OUTLOOK_CALENDAR);
            if (calendar_exists)
            {
                string calendarId = (string)ApplicationData.Current.LocalSettings.Values[OUTLOOK_CALENDAR];
                this.appCalendar = await appointmentStore.GetAppointmentCalendarAsync(calendarId);
            }
            else
            {
                CreateCalendar();
            }
        }

        public async Task HandleAppointments()
        {
            while (true)
            {
                CalendarAppointment calendarAppointment = appointmentCollection.Take();
                Appointment appointment = calendarAppointment.ToAppointment();
                switch (calendarAppointment.Action)
                {
                    case AppointmentAction.AddItem:
                        await AddAppointment(appointment);
                        break;
                    case AppointmentAction.RemoveItem:
                        await RemoveAppointment(appointment);
                        break;
                    case AppointmentAction.ChangeItem:
                        await ModifyAppointment(appointment);
                        break;
                    default:
                        throw new NotImplementedException();
                }
            }
        }

        async void CreateCalendar()
        {
            this.appCalendar = await appointmentStore.CreateAppointmentCalendarAsync(OUTLOOK_CALENDAR);
            ApplicationData.Current.LocalSettings.Values[OUTLOOK_CALENDAR] = this.appCalendar.LocalId;
        }

        public async Task AddAppointment(Appointment appointment)
        {
            await appCalendar.SaveAppointmentAsync(appointment);
            calendarStorageSettings.AddLocalIdMapping(appointment.RoamingId, appointment.LocalId);
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
            calendarStorageSettings.AddLocalIdMapping(appointment.RoamingId, appointment.LocalId);
        }

        public async Task<List<Appointment>> GetEventsOnDate(DateTime date)
        {
            this.appointmentStore = await AppointmentManager.RequestStoreAsync(AppointmentStoreAccessType.AllCalendarsReadWrite);
            this.appCalendar = await appointmentStore.CreateAppointmentCalendarAsync(OUTLOOK_CALENDAR);
            IReadOnlyList<Appointment> appointments = await appointmentStore.FindAppointmentsAsync(new DateTimeOffset(date.Date), TimeSpan.FromHours(24));
            return new List<Appointment>(appointments);
        }
    }
}
