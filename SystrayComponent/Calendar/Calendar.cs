﻿using Microsoft.Office.Interop.Outlook;
using SystrayComponent.AppointmentDetails;

namespace SystrayComponent.Calendar
{
    internal class Calendar
    {
        const string OUTLOOK_NAMESPACE = "MAPI";
        Microsoft.Office.Interop.Outlook.Application outlookApp;
        NameSpace mapiNamespace;
        MAPIFolder calendarFolder;
        Items calendarItems;
        List<AppointmentItem> calendarItemsList;
        AppointmentSender appointmentSender;
        ItemsEvents_ItemAddEventHandler addEvent;
        ItemsEvents_ItemChangeEventHandler changeEvent;
        ItemEvents_10_BeforeDeleteEventHandler beforeDeleteEvent;

        public Calendar(ItemsEvents_ItemAddEventHandler addEvent, ItemsEvents_ItemChangeEventHandler changeEvent, ItemEvents_10_BeforeDeleteEventHandler beforeDeleteEvent)
        {
            appointmentSender = new AppointmentSender();
            outlookApp = new Microsoft.Office.Interop.Outlook.Application();
            mapiNamespace = outlookApp.GetNamespace(OUTLOOK_NAMESPACE);
            calendarFolder = mapiNamespace.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

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
            calendarItems.ItemAdd += this.addEvent;
            calendarItems.ItemChange += this.changeEvent;

            foreach (AppointmentItem item in calendarItems)
            {
                item.BeforeDelete += this.beforeDeleteEvent;
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
