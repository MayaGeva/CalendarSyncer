using System.Windows.Forms;
using Windows.ApplicationModel;
using Windows.ApplicationModel.AppService;
using Windows.ApplicationModel.Core;
using Windows.Foundation.Collections;
using Windows.Storage;

namespace SystrayComponent
{
    public partial class Form1 : Form
    {
        private AppServiceConnection? connection = null;
        private NotifyIcon notifyIcon;
        public Form1()
        {
            InitializeComponent();
            ToolStripMenuItem openMenuItem = new ToolStripMenuItem("Open App", null, new EventHandler(OpenApp));
            ToolStripMenuItem exitMenuItem = new ToolStripMenuItem("Exit", null, new EventHandler(Exit));

            notifyIcon = new NotifyIcon();
            notifyIcon.Icon = Resources.CalendarIcon;
            notifyIcon.DoubleClick += new EventHandler(OpenApp);
            var menuStrip = new ContextMenuStrip();
            menuStrip.Items.Add(openMenuItem);
            menuStrip.Items.Add(exitMenuItem);
            notifyIcon.ContextMenuStrip = menuStrip;
            notifyIcon.Visible = true;
            this.ShowInTaskbar = false;
            Hide();
        }
        
        private async void OpenApp(object? sender, EventArgs e)
        {
            IEnumerable<AppListEntry> appListEntries = await Package.Current.GetAppListEntriesAsync();
            await appListEntries.First().LaunchAsync();
        }

        private async void Exit(object? sender, EventArgs e)
        {
            ValueSet message = new ValueSet
            {
                { "exit", "" }
            };
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
            if (connection != null)
            {
                connection.ServiceClosed -= Connection_ServiceClosed;
                connection = null;
            }
        }
    }
}