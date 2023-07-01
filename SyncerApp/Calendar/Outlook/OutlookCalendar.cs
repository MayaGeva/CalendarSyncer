using Microsoft.Office.Interop.Outlook;
using System.Globalization;

namespace SyncerApp.Calendar.Outlook
{
    internal class OutlookCalendar
    {
        const string OUTLOOK_NAMESPACE = "MAPI";
        Microsoft.Office.Interop.Outlook.Application outlookApp;
        NameSpace mapiNamespace;
        MAPIFolder calendarFolder;
        Items calendarItems;
        public Recipient CurrentUser { get; }
        List<AppointmentItem> calendarItemsList;
        ItemsEvents_ItemAddEventHandler addEvent;
        ItemsEvents_ItemChangeEventHandler changeEvent;
        ItemEvents_10_BeforeDeleteEventHandler beforeDeleteEvent;

        public OutlookCalendar(ItemsEvents_ItemAddEventHandler addEvent, ItemsEvents_ItemChangeEventHandler changeEvent, ItemEvents_10_BeforeDeleteEventHandler beforeDeleteEvent)
        {
            outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            mapiNamespace = outlookApp.GetNamespace(OUTLOOK_NAMESPACE);
            calendarFolder = mapiNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

            CurrentUser = mapiNamespace.CurrentUser;

            calendarItems = calendarFolder.Items;
            calendarItems.IncludeRecurrences = true;

            this.addEvent = addEvent;
            this.changeEvent = changeEvent;
            this.beforeDeleteEvent = beforeDeleteEvent;
            calendarItemsList = new List<AppointmentItem>();
            SetupEventHandlers();
        }
        void SetupEventHandlers()
        {
            calendarItems.ItemAdd += addEvent;
            calendarItems.ItemChange += changeEvent;

            foreach (AppointmentItem item in calendarItems)
            {
                item.BeforeDelete += beforeDeleteEvent;
                calendarItemsList.Add(item);
            }
        }

        public void AddItem(AppointmentItem appointmentItem)
        {
            appointmentItem.BeforeDelete += beforeDeleteEvent;
            calendarItemsList.Add(appointmentItem);
        }

        public void RemoveItem(AppointmentItem appointmentItem)
        {
            appointmentItem.BeforeDelete -= beforeDeleteEvent;
            calendarItemsList.Remove(appointmentItem);
        }

        public List<AppointmentItem> GetAllAppointments()
        {
            List<AppointmentItem> appointments = new List<AppointmentItem>();
            Items outlookCalendarItems = calendarFolder.Items;
            foreach (AppointmentItem item in outlookCalendarItems)
            {
                appointments.Add(item);
            }
            return appointments;
        }

        public List<AppointmentItem> GetCalendarAppointmentsModifiedAfter(DateTime time)
        {
            string timeFormat = CultureInfo.CurrentCulture.DateTimeFormat.ShortDatePattern;
            string sfilter = string.Format("[ModifiedTime]>={0}", time.ToString(timeFormat));
            return GetAppointmentsFiltered(sfilter);
        }

        /// <summary>
        /// Get all appointments with a certain filter
        /// </summary>
        /// <param name="filter">The filter to be use on the appointments</param>
        /// <returns>A list of appointments that match the filter</returns>
        List<AppointmentItem> GetAppointmentsFiltered(string filter)
        {
            List<AppointmentItem> appointments = new List<AppointmentItem>();
            Items outlookCalendarItems = calendarFolder.Items.Restrict(filter);
            outlookCalendarItems.IncludeRecurrences = true;
            foreach (AppointmentItem item in outlookCalendarItems)
            {
                appointments.Add(item);
            }
            return appointments;
        }
    }
}
