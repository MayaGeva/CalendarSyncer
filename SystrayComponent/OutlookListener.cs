using Microsoft.Office.Interop.Outlook;
using System.Collections.Concurrent;
using SystrayComponent.AppointmentDetails;
using SystrayComponent.Calendar;
using Windows.Storage;

namespace SystrayComponent
{
    internal class OutlookListener
    {
        Calendar.Calendar calendar;
        BlockingCollection<CalendarAppointment> calendarAppointments;
        public OutlookListener(BlockingCollection<CalendarAppointment> calendarAppointments)
        {
            this.calendar = new Calendar.Calendar(OutlookItemAdd, OutlookItemChange, ItemBeforeDelete);
            this.calendarAppointments = calendarAppointments;
        }

        void FirstTimeRun(BlockingCollection<CalendarAppointment> appointmentQueue)
        {
            bool firstTime = ApplicationData.Current.LocalSettings.Values.ContainsKey("lastSyncTime");
            if (firstTime)
            {
                List<CalendarAppointment> appointments = calendar.GetAllAppointments();
                foreach (CalendarAppointment appointment in appointments)
                {
                    appointmentQueue.Add(appointment);
                }
                ApplicationData.Current.LocalSettings.Values["lastSyncTime"] = DateTime.Now.ToString();
            }
            else
            {
                DateTime lastUpdate = DateTime.Parse((string)ApplicationData.Current.LocalSettings.Values["lastSyncTime"]);
                List<CalendarAppointment> appointments = calendar.GetCalendarAppointmentsModifiedAfter(lastUpdate);
                foreach (CalendarAppointment appointment in appointments)
                {
                    appointmentQueue.Add(appointment);
                }
            }
        }

        private void ItemBeforeDelete(object Item, ref bool Cancel)
        {
            Console.WriteLine("An Item is being deleted!");
            if (!Cancel && Item is AppointmentItem item)
            {
                calendar.RemoveItem(item);
                CalendarAppointment appointment = new CalendarAppointment(item, AppointmentAction.RemoveItem);
                this.calendarAppointments.Add(appointment);
            }
        }

        private void OutlookItemChange(object Item)
        {
            Console.WriteLine("An Item has changed!");
            if (Item is AppointmentItem item)
            {
                CalendarAppointment appointment = new CalendarAppointment(item, AppointmentAction.ChangeItem);
                this.calendarAppointments.Add(appointment);
            }
        }

        private void OutlookItemAdd(object Item)
        {
            Console.WriteLine("An Item has been added!");

            if (Item is AppointmentItem item)
            {
                calendar.AddItem(item);
                CalendarAppointment appointment = new CalendarAppointment(item, AppointmentAction.AddItem);
                this.calendarAppointments.Add(appointment);
            }
        }
    }
}
