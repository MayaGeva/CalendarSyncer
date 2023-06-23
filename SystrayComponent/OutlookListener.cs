using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SystrayComponent.AppointmentDetails;
using SystrayComponent.Calendar;

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
