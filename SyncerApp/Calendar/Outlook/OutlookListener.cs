using Microsoft.Office.Interop.Outlook;
using System.Collections.Concurrent;
using Windows.Storage;

namespace SyncerApp.Calendar.Outlook
{
    internal class OutlookListener
    {
        public OutlookCalendar Calendar { get; }
        BlockingCollection<CalendarAppointment> calendarAppointments;
        const string LAST_SYNC_TIME = "lastSyncTime";
        public OutlookListener(BlockingCollection<CalendarAppointment> calendarAppointments)
        {
            Calendar = new OutlookCalendar(OutlookItemAdd, OutlookItemChange, ItemBeforeDelete);
            this.calendarAppointments = calendarAppointments;
        }

        public void FirstTimeRun()
        {
            bool firstTime = !ApplicationData.Current.LocalSettings.Values.ContainsKey(LAST_SYNC_TIME);
            List<AppointmentItem> appointments;
            if (firstTime)
            {
                appointments = Calendar.GetAllAppointments();
                ApplicationData.Current.LocalSettings.Values[LAST_SYNC_TIME] = DateTime.Now.ToString();
            }
            else
            {
                DateTime lastUpdate = DateTime.Parse((string)ApplicationData.Current.LocalSettings.Values[LAST_SYNC_TIME]);
                appointments = Calendar.GetCalendarAppointmentsModifiedAfter(lastUpdate);
            }
            foreach (AppointmentItem appointment in appointments)
            {
                CalendarAppointment calendarAppointment = new(AppointmentAction.AddItem, appointment);
                AddAppointmentToQueue(calendarAppointment);
            }
        }

        void AddAppointmentToQueue(CalendarAppointment appointment)
        {
            Mutex mutex = new(true, "AppointmentsQueue", out bool createdNew);
            if (createdNew)
            {
                calendarAppointments.Add(appointment);
            }
            else
            {
                mutex.WaitOne();
                calendarAppointments.Add(appointment);
            }
            mutex.ReleaseMutex();
        }

        private void ItemBeforeDelete(object Item, ref bool Cancel)
        {
            Console.WriteLine("An Item is being deleted!");
            if (!Cancel && Item is AppointmentItem item)
            {
                Calendar.RemoveItem(item);
                CalendarAppointment appointment = new(AppointmentAction.RemoveItem, item);
                AddAppointmentToQueue(appointment);
            }
        }

        private void OutlookItemChange(object Item)
        {
            Console.WriteLine("An Item has changed!");
            if (Item is AppointmentItem item)
            {
                CalendarAppointment appointment = new(AppointmentAction.ChangeItem, item);
                AddAppointmentToQueue(appointment);
            }
        }

        private void OutlookItemAdd(object Item)
        {
            Console.WriteLine("An Item has been added!");

            if (Item is AppointmentItem item)
            {
                Calendar.AddItem(item);
                CalendarAppointment appointment = new(AppointmentAction.AddItem, item);
                AddAppointmentToQueue(appointment);
            }
        }
    }
}
