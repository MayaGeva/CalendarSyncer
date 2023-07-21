using Microsoft.Win32;
using SyncerApp.Calendar;
using SyncerApp.Calendar.Outlook;
using SyncerApp.Calendar.Windows;
using System.Collections.Concurrent;
using Windows.Storage;

namespace SyncerApp
{
    internal static class Program
    {
        const string START_KEY = @"SOFTWARE\Microsoft\Windows\CurrentVersion\Run";
        const string OUTLOOK_LAUNCHER = @"\OutlookLauncher.exe";
        const string PROTOCOL_REGISTRY = @"ms-outlook\shell\open\command";
        /// <summary>
        ///  The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // To customize application configuration such as set high DPI settings or default font,
            // see https://aka.ms/applicationconfiguration.
            ApplicationConfiguration.Initialize();

            //AddToStartUp();
            //AddProtocolEntry();
            WindowsCalendar windowsCalendar = new();

            BlockingCollection<CalendarAppointment> calendarAppointments = new();
            OutlookListener outlookListener = new(calendarAppointments);
            outlookListener.FirstTimeRun();
            
            AppointmentConverter appointmentConverter = new(outlookListener.Calendar.CurrentUser);
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

        static void AddProtocolEntry()
        {
            string packagePath = Windows.ApplicationModel.Package.Current.InstalledLocation.Path;
            string launcherPath = Path.Join(packagePath, OUTLOOK_LAUNCHER);
            
            string[] RegKeys = PROTOCOL_REGISTRY.Split('\\');
            
            RegistryKey regKey = Registry.ClassesRoot;
            foreach (string key in RegKeys)
            {
                regKey = regKey.CreateSubKey(key);
            }
            regKey.SetValue("Default", $"\"{launcherPath}\" \"%1\"");
        }
    }
}