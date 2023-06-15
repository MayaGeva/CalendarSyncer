using Microsoft.Office.Interop.Outlook;
using OutlookListener.AppointmentDetails;

namespace OutlookListener.Calendar
{
    internal class Calendar
    {
        const string OUTLOOK_NAMESPACE = "MAPI";
        Application outlookApp;
        NameSpace mapiNamespace;
        MAPIFolder calendarFolder;
        Items calendarItems;
        AppointmentSender appointmentSender;

        public Calendar()
        {
            appointmentSender = new AppointmentSender();
            outlookApp = new Application();
            mapiNamespace = outlookApp.GetNamespace(OUTLOOK_NAMESPACE);
            calendarFolder = mapiNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
            calendarItems = calendarFolder.Items;
            calendarItems.IncludeRecurrences = true;
            SetupEventHandlers();
        }
        void SetupEventHandlers()
        {
            calendarItems.ItemAdd += OutlookListenerItems_ItemAdd;
            calendarItems.ItemChange += OutlookListenerItems_ItemChange;

            foreach (AppointmentItem item in calendarItems)
            {
                item.BeforeDelete += CalendarItem_BeforeDelete;
            }
        }

        private void CalendarItem_BeforeDelete(object Item, ref bool Cancel)
        {
            Console.WriteLine("An Item is being deleted!");
            if (!Cancel && Item is AppointmentItem item)
            {
                CalendarAppointment appointment = new CalendarAppointment(item, AppointmentAction.RemoveItem);
                this.appointmentSender.SendAppointment(appointment);
                item.BeforeDelete -= CalendarItem_BeforeDelete;

            }
        }

        private void OutlookListenerItems_ItemChange(object Item)
        {
            Console.WriteLine("An Item has changed!");
            if (Item is AppointmentItem item)
            {
                CalendarAppointment appointment = new CalendarAppointment(item, AppointmentAction.ChangeItem);
                this.appointmentSender.SendAppointment(appointment);
            }
        }

        private void OutlookListenerItems_ItemAdd(object Item)
        {
            Console.WriteLine("An Item has been added!");

            if (Item is AppointmentItem item)
            {
                // get appointment xml
                item.BeforeDelete += CalendarItem_BeforeDelete;
                calendarItems.Add(item);
                CalendarAppointment appointment = new CalendarAppointment(item, AppointmentAction.AddItem);
                this.appointmentSender.SendAppointment(appointment);
            }
        }


        /// <summary>
        /// Get all appointments with a certain filter
        /// </summary>
        /// <param name="filter">The filter to be use on the appointments</param>
        /// <returns>A list of appointments that match the filter</returns>
        List<CalendarAppointment> GetAppointmentsFiltered(string filter)
        {
            List<CalendarAppointment> appointments = new List<CalendarAppointment>();
            Items outlookCalendarItems = calendarFolder.Items.Restrict(filter);
            outlookCalendarItems.IncludeRecurrences = true;
            foreach (AppointmentItem item in outlookCalendarItems)
            {
                appointments.Add(new CalendarAppointment(item));
            }
            return appointments;
        }
    }
}
