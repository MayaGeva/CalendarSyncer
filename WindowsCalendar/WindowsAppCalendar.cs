using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Windows.ApplicationModel.Appointments;

namespace WindowsCalendar.Calendar
{
    internal class WindowsAppCalendar
    {
        AppointmentStore appointmentStore;
        AppointmentCalendar appCalendar;
        public WindowsAppCalendar() 
        {
            InitCalendar();
        }
        async void InitCalendar()
        {
            this.appointmentStore = await AppointmentManager.RequestStoreAsync(AppointmentStoreAccessType.AllCalendarsReadWrite);
            this.appCalendar = await appointmentStore.CreateAppointmentCalendarAsync("OutlookCalendar");
        }
        public async Task AddAppointment(Appointment appointment)
        {
            await appCalendar.SaveAppointmentAsync(appointment);
            
        }
        public async Task RemoveAppointment(Appointment appointment) 
        {
            await appCalendar.DeleteAppointmentAsync(appointment.LocalId);
        }
        public async Task ModifyAppointment(Appointment appointment)
        {

        }
        public async Task<List<Appointment>> GetEventsOnDate(DateTime date)
        {
            this.appointmentStore = await AppointmentManager.RequestStoreAsync(AppointmentStoreAccessType.AllCalendarsReadWrite);
            this.appCalendar = await appointmentStore.CreateAppointmentCalendarAsync("OutlookListener");
            IReadOnlyList<Appointment> appointments = await appointmentStore.FindAppointmentsAsync(new DateTimeOffset(date.Date), TimeSpan.FromHours(24));
            return new List<Appointment>(appointments);
        }
    }
}
