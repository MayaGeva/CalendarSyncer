namespace SyncerApp
{
    public partial class SystrayComponent : Form
    {
        NotifyIcon notifyIcon;
        public SystrayComponent()
        {
            InitializeComponent();
            ToolStripMenuItem exitMenuItem = new("Exit", null, new EventHandler(Exit));
            var menuStrip = new ContextMenuStrip();
            menuStrip.Items.Add(exitMenuItem);

            notifyIcon = new()
            {
                Icon = Resources.SystrayIcon,
                ContextMenuStrip = menuStrip
            };
            notifyIcon.Visible = true;
            this.ShowInTaskbar = false;
            this.Visible = false;
        }
        private void Exit(object? sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
