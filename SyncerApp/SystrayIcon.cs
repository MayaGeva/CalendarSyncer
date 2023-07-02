using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SyncerApp
{
    internal class SystrayIcon : ApplicationContext
    {
        NotifyIcon notifyIcon;
        public SystrayIcon()
        {
            ToolStripMenuItem exitMenuItem = new("Exit", null, new EventHandler(Exit));
            var menuStrip = new ContextMenuStrip();
            menuStrip.Items.Add(exitMenuItem);

            notifyIcon = new()
            {
                Icon = Resources.SystrayIcon,
                ContextMenuStrip = menuStrip,
                Visible = true
            };
        }
        private void Exit(object? sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
