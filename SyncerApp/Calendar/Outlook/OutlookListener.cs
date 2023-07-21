using Microsoft.Office.Interop.Outlook;
using System.Collections.Concurrent;
using Windows.Storage;

namespace SyncerApp.Calendar.Outlook
{
    internal class OutlookListener
    {
        public OutlookCalendar Calendar { get; }
        BlockingCollection<CalendarAppointment> calendarAppointments;
        const string LAST_SYNC_TIME = "LastSyncTime";

        public OutlookListener(BlockingCollection<CalendarAppointment> calendarAppointments)
        {
            Calendar = new OutlookCalendar(OutlookItemAdd, OutlookItemChange, ItemBeforeDelete);
            this.calendarAppointments = calendarAppointments;
        }

        /// <summary>
        /// This is called when the app is first opened
        /// </summary>
        public void FirstTimeRun()
        {
            bool firstTime = !ApplicationData.Current.LocalSettings.Values.ContainsKey(LAST_SYNC_TIME); // True if the calendar's never been synced
            List<AppointmentItem> appointments = new();

            if (firstTime) {
                appointments = Calendar.GetAllAppointments();
            } else {
                DateTime lastUpdate = DateTime.Parse((string)ApplicationData.Current.LocalSettings.Values[LAST_SYNC_TIME]);
                appointments = Calendar.GetCalendarAppointmentsModifiedAfter(lastUpdate);
            }

            foreach (AppointmentItem appointment in appointments) {
                calendarAppointments.Add(new(AppointmentAction.AddItem, appointment));
            }
            calendarAppointments.Add(new(AppointmentAction.NoAction, null));
        }

        /// <summary>
        /// This is called before an item is deleted from the outlook calendar folder
        /// </summary>
        /// <param name="Item">The item that is being deleted</param>
        /// <param name="Cancel">Whether or not the user pressed cancel</param>
        private void ItemBeforeDelete(object Item, ref bool Cancel)
        {
            Console.WriteLine("An Item is being deleted!");
            if (!Cancel && Item is AppointmentItem item) {
                Calendar.RemoveItem(item);
                CalendarAppointment appointment = new(AppointmentAction.RemoveItem, item);
                calendarAppointments.Add(appointment);
            } else if (!Cancel && Item is MeetingItem meeting) {
                var appointment = meeting.GetAssociatedAppointment(true);
                Calendar.AddItem(appointment);
                CalendarAppointment calAppointment = new(AppointmentAction.ChangeItem, appointment);
                calendarAppointments.Add(calAppointment);
            }
        }

        /// <summary>
        /// This is called when an item is changed in the outlook calendar folder
        /// </summary>
        /// <param name="Item">The item that was changed</param>
        private void OutlookItemChange(object Item)
        {
            Console.WriteLine("An Item has changed!");
            if (Item is AppointmentItem item) {
                CalendarAppointment appointment = new(AppointmentAction.ChangeItem, item);
                calendarAppointments.Add(appointment);
            } else if (Item is MeetingItem meeting) {
                var appointment = meeting.GetAssociatedAppointment(true);
                Calendar.AddItem(appointment);
                CalendarAppointment calAppointment = new(AppointmentAction.ChangeItem, appointment);
                calendarAppointments.Add(calAppointment);
            }
        }

        /// <summary>
        /// This is called when an item is added to the outlook calendar folder
        /// </summary>
        /// <param name="Item">The outlook item added to the folder</param>
        private void OutlookItemAdd(object Item)
        {
            Console.WriteLine("An Item has been added!");

            if (Item is AppointmentItem item) {
                Calendar.AddItem(item);
                CalendarAppointment appointment = new(AppointmentAction.AddItem, item);
                calendarAppointments.Add(appointment);
            } else if (Item is MeetingItem meeting) {
                var appointment = meeting.GetAssociatedAppointment(true);
                Calendar.AddItem(appointment);
                CalendarAppointment calAppointment = new(AppointmentAction.AddItem, appointment);
                calendarAppointments.Add(calAppointment);
            }
        }
    }
}
