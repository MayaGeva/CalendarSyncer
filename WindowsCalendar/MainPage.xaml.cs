using System;
using System.Collections.Concurrent;
using System.Threading.Tasks;
using Windows.ApplicationModel;
using Windows.Foundation.Metadata;
using Windows.Foundation;
using Windows.Security.Authentication.Web;
using Windows.Storage;
using Windows.UI.Core.Preview;
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
        const string PACKAGE_SID = "PackageSid";
        public MainPage()
        {
            this.InitializeComponent();
            
            SystemNavigationManagerPreview mgr = SystemNavigationManagerPreview.GetForCurrentView();
            mgr.CloseRequested += SystemNavigationManager_CloseRequested;
            StartOutlookListener();
        }

        async void StartOutlookListener()
        {
            ApplicationData.Current.LocalSettings.Values[PACKAGE_SID] = WebAuthenticationBroker.GetCurrentApplicationCallbackUri().Host.ToUpper();
            await FullTrustProcessLauncher.LaunchFullTrustProcessForCurrentAppAsync();
        }

        private async void SystemNavigationManager_CloseRequested(object sender, SystemNavigationCloseRequestedPreviewEventArgs e)
        {
            Deferral deferral = e.GetDeferral();
            
            if (ApiInformation.IsApiContractPresent(
                    "Windows.ApplicationModel.FullTrustAppContract", 1, 0))
            {
                await FullTrustProcessLauncher.LaunchFullTrustProcessForCurrentAppAsync();
            }
            e.Handled = false;
            deferral.Complete();
        }
    }
}
