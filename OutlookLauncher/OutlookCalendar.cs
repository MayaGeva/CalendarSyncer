using Microsoft.Office.Interop.Outlook;
using System.Globalization;

namespace OutlookLauncher
{
    internal class OutlookCalendar
    {
        const string OUTLOOK_NAMESPACE = "MAPI";
        Microsoft.Office.Interop.Outlook.Application outlookApp;
        NameSpace mapiNamespace;
        MAPIFolder calendarFolder;
        Items calendarItems;
        List<AppointmentItem> calendarItemsList;

        public OutlookCalendar()
        {
            outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            mapiNamespace = outlookApp.GetNamespace(OUTLOOK_NAMESPACE);
            calendarFolder = mapiNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

            calendarItems = calendarFolder.Items;
            calendarItems.IncludeRecurrences = true;

            calendarItemsList = new();
            foreach (AppointmentItem item in calendarItems) {
                calendarItemsList.Add(item);
            }
        }

        public AppointmentItem? GetAppointmentFromGlobalId(string entryId)
        {
            return calendarItemsList.Where(i => i.GlobalAppointmentID == entryId).FirstOrDefault();
        }

        /// <summary>
        /// Get all appointments with a certain filter
        /// </summary>
        /// <param name="filter">The filter to be use on the appointments</param>
        /// <returns>A list of appointments that match the filter</returns>
        List<AppointmentItem> GetAppointmentsFiltered(string filter)
        {
            List<AppointmentItem> appointments = new();
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
