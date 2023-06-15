using System.Windows.Forms;

namespace SystrayComponent
{
    public partial class Form1 : Form
    {
        NotifyIcon notifyIcon;
        public Form1()
        {
            InitializeComponent();
            MenuItem openMenuItem = new MenuItem("Open UWP", new EventHandler(OpenApp));
            MenuItem sendMenuItem =
            new MenuItem("Send message to UWP", new EventHandler(SendToUWP));
            MenuItem legacyMenuItem =
            new MenuItem("Open legacy companion", new EventHandler(OpenLegacy));
            MenuItem exitMenuItem = new MenuItem("Exit", new EventHandler(Exit));
            openMenuItem.DefaultItem = true;

            notifyIcon = new NotifyIcon();
            notifyIcon.DoubleClick += new EventHandler(OpenApp);
            notifyIcon.Icon = SystrayComponent.Properties.Resources.Icon1;
            notifyIcon.ContextMenu = new ContextMenu(new MenuItem[]
              { openMenuItem, sendMenuItem, legacyMenuItem, exitMenuItem });
            notifyIcon.Visible = true;
        }
    }
}