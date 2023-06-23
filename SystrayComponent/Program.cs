using System.Collections.Concurrent;
using SystrayComponent.Calendar;

namespace SystrayComponent
{
    internal static class Program
    {
        const string MUTEX_NAME = "MySystrayExtensionMutex";

        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.

            Mutex? mutex;
            if (!Mutex.TryOpenExisting(MUTEX_NAME, out mutex))
            {
                mutex = new Mutex(false, MUTEX_NAME);
                ApplicationConfiguration.Initialize();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);

                BlockingCollection<CalendarAppointment> calendarAppointments = new BlockingCollection<CalendarAppointment>();
                OutlookListener outlookListener = new OutlookListener(calendarAppointments);
                AppointmentSender appointmentSender = new AppointmentSender(calendarAppointments);
                Task _ = Task.Run(() => appointmentSender.HandleAppointments());
                Application.Run(new Form1());
                mutex.Close();
            }
        }
    }
}