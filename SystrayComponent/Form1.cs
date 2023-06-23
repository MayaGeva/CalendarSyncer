using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using SystrayComponent.AppointmentDetails;
using SystrayComponent.Calendar;
using Windows.ApplicationModel;
using Windows.ApplicationModel.AppService;
using Windows.ApplicationModel.Core;
using Windows.Foundation.Collections;

namespace SystrayComponent
{
    public partial class Form1 : Form
    {
        private AppServiceConnection? connection = null;
        private NotifyIcon notifyIcon;
        Calendar.Calendar calendar;
        AppointmentSender appointmentSender;
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

            appointmentSender = new AppointmentSender();
            calendar = new Calendar.Calendar(OutlookItemAdd, OutlookItemChange, ItemBeforeDelete);
        }

        private void ItemBeforeDelete(object Item, ref bool Cancel)
        {
            Console.WriteLine("An Item is being deleted!");
            if (!Cancel && Item is AppointmentItem item)
            {
                CalendarAppointment appointment = new CalendarAppointment(item, AppointmentAction.RemoveItem);
                this.appointmentSender.SendAppointment(appointment);
                calendar.RemoveItem(item);
            }
        }

        private void OutlookItemChange(object Item)
        {
            Console.WriteLine("An Item has changed!");
            if (Item is AppointmentItem item)
            {
                CalendarAppointment appointment = new CalendarAppointment(item, AppointmentAction.ChangeItem);
                this.appointmentSender.SendAppointment(appointment);
            }
        }

        private void OutlookItemAdd(object Item)
        {
            Console.WriteLine("An Item has been added!");

            if (Item is AppointmentItem item)
            {
                // get appointment xml
                item.BeforeDelete += ItemBeforeDelete;
                calendar.AddItem(item);
                CalendarAppointment appointment = new CalendarAppointment(item, AppointmentAction.AddItem);
                this.appointmentSender.SendAppointment(appointment);
            }
        }


        private async void OpenApp(object? sender, EventArgs e)
        {
            IEnumerable<AppListEntry> appListEntries = await Package.Current.GetAppListEntriesAsync();
            await appListEntries.First().LaunchAsync();
        }

        private async void Exit(object sender, EventArgs e)
        {
            ValueSet message = new ValueSet();
            message.Add("exit", "");
            await SendToUWP(message);
            System.Windows.Forms.Application.Exit();
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