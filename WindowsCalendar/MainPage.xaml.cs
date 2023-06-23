using System;
using System.Collections.Concurrent;
using System.Threading.Tasks;
using Windows.ApplicationModel;
using Windows.Security.Authentication.Web;
using Windows.Storage;
using Windows.UI.Xaml.Controls;
using WindowsCalendar;
using WindowsCalendar.AppointmentDetails;
using WindowsCalendar.Calendar;

// The Blank Page item template is documented at https://go.microsoft.com/fwlink/?LinkId=402352&clcid=0x409

namespace CalendarSyncer
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            this.InitializeComponent();
            BlockingCollection<CalendarAppointment> calendarAppointments = new BlockingCollection<CalendarAppointment>();
            Task pipeServer = Task.Run(() => new PipeServer(calendarAppointments).RunServer());
            Task appointmentHandler = Task.Run(() => new WindowsAppCalendar(calendarAppointments).HandleAppointments());
            StartOutlookListener();
        }

        async void StartOutlookListener()
        {
            ApplicationData.Current.LocalSettings.Values["PackageSid"] = WebAuthenticationBroker.GetCurrentApplicationCallbackUri().Host.ToUpper();
            await FullTrustProcessLauncher.LaunchFullTrustProcessForCurrentAppAsync();
        }
    }
}
