using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Threading.Tasks;
using Windows.ApplicationModel;
using Windows.ApplicationModel.Core;
using Windows.ApplicationModel.AppService;
using Windows.Foundation.Collections;

namespace SystrayComponent
{
    class SystrayApplicationContext : ApplicationContext
    {
        private AppServiceConnection connection = null;
        private NotifyIcon notifyIcon = null;
        Calendar.Calendar calendar;

        public SystrayApplicationContext()
        {
            ToolStripMenuItem openMenuItem = new ToolStripMenuItem("Open App", null, new EventHandler(OpenApp));
            ToolStripMenuItem exitMenuItem = new ToolStripMenuItem("Exit", null, new EventHandler(Exit));

            notifyIcon = new NotifyIcon();
            notifyIcon.DoubleClick += new EventHandler(OpenApp);
            var menuStrip = new ContextMenuStrip();
            menuStrip.Items.Add(openMenuItem);
            menuStrip.Items.Add(exitMenuItem);
            notifyIcon.ContextMenuStrip = menuStrip;
            notifyIcon.Visible = true;
            calendar = new Calendar.Calendar();
        }

        private async void OpenApp(object sender, EventArgs e)
        {
            IEnumerable<AppListEntry> appListEntries = await Package.Current.GetAppListEntriesAsync();
            await appListEntries.First().LaunchAsync();
        }

        private async void Exit(object sender, EventArgs e)
        {
            ValueSet message = new ValueSet();
            message.Add("exit", "");
            await SendToUWP(message);
            Application.Exit();
        }

        private async Task SendToUWP(ValueSet message)
        {
            if (connection == null)
            {
                connection = new AppServiceConnection();
                connection.PackageFamilyName = Package.Current.Id.FamilyName;
                connection.AppServiceName = "SystrayExtensionService";
                connection.ServiceClosed += Connection_ServiceClosed;
                AppServiceConnectionStatus connectionStatus = await connection.OpenAsync();
                if (connectionStatus != AppServiceConnectionStatus.Success)
                {
                    MessageBox.Show("Status: " + connectionStatus.ToString());
                    return;
                }
            }

            await connection.SendMessageAsync(message);
        }

        private void Connection_ServiceClosed(AppServiceConnection sender, AppServiceClosedEventArgs args)
        {
            connection.ServiceClosed -= Connection_ServiceClosed;
            connection = null;
        }
    }
}