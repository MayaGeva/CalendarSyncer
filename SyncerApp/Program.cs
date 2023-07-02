using Microsoft.Win32;
using SyncerApp.Calendar;
using SyncerApp.Calendar.Outlook;
using SyncerApp.Calendar.Windows;
using System.Collections.Concurrent;

namespace SyncerApp
{
    internal static class Program
    {
        const string START_KEY = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();
            
            AddToStartUp();

            BlockingCollection<CalendarAppointment> calendarAppointments = new();
            OutlookListener outlookListener = new(calendarAppointments);
            outlookListener.FirstTimeRun();
            
            AppointmentConverter appointmentConverter = new(outlookListener.Calendar.CurrentUser);
            CalendarStorageSettings storageSettings = new();
            WindowsCalendar windowsCalendar = new(storageSettings);
            WindowsCalendarSyncer windowsCalendarSyncer = new(calendarAppointments, appointmentConverter, windowsCalendar);

            CancellationTokenSource cancellationToken = new();
            Task handleAppointments = Task.Run(() => windowsCalendarSyncer.HandleAppointments(cancellationToken.Token));
            Application.ApplicationExit += (s, e) => cancellationToken.Cancel();
            
            Application.Run(new SystrayIcon());
        }
        static void AddToStartUp()
        {
            RegistryKey? startup = Registry.CurrentUser.OpenSubKey(START_KEY, true);
            startup?.SetValue("CalendarSyncer", Application.ExecutablePath);
        }
    }
}