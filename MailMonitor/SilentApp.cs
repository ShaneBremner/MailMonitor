using System;
using System.Drawing;
using System.Windows.Forms;

namespace MailMonitor
{
    public class SilentApp : ApplicationContext
    {
        private readonly System.Windows.Forms.Timer _heartbeatTimer;
        private readonly NotifyIcon _trayIcon;
        private readonly OutlookWatcherTask _watcher;

        public SilentApp()
        {
            _trayIcon = new NotifyIcon
            {
                Icon = SystemIcons.Application,
                Visible = true,
                Text = "Mail Monitor",
                ContextMenuStrip = BuildTrayMenu()
            };

            _heartbeatTimer = new System.Windows.Forms.Timer();
            ApplyHeartbeatInterval();
            _heartbeatTimer.Tick += OnHeartbeat;
            _heartbeatTimer.Start();

            // Start the inbox watcher immediately on startup.
            _watcher = new OutlookWatcherTask();
            _watcher.Start();
        }

        private ContextMenuStrip BuildTrayMenu()
        {
            var menu = new ContextMenuStrip();
            menu.Items.Add("Check Now", null, (s, e) => OnHeartbeat(s, e));
            menu.Items.Add("Set Interval (minutes)...", null, OnSetInterval);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Exit", null, OnExit);
            return menu;
        }

        private void OnHeartbeat(object? sender, EventArgs e)
        {
            OutlookTask.CheckUnreadMail();
            // Add more heartbeat tasks here as needed.
        }

        private void OnSetInterval(object? sender, EventArgs e)
        {
            string input = Microsoft.VisualBasic.Interaction.InputBox(
                $"Enter heartbeat interval in minutes (current: {SettingsService.HeartbeatMinutes}):",
                "Mail Monitor – Set Interval",
                SettingsService.HeartbeatMinutes.ToString()
            );

            if (int.TryParse(input, out int minutes) && minutes > 0)
            {
                SettingsService.HeartbeatMinutes = minutes;
                ApplyHeartbeatInterval();
            }
        }

        private void ApplyHeartbeatInterval()
        {
            _heartbeatTimer.Stop();
            _heartbeatTimer.Interval = (int)TimeSpan.FromMinutes(SettingsService.HeartbeatMinutes).TotalMilliseconds;
            _heartbeatTimer.Start();
        }

        private void OnExit(object? sender, EventArgs e)
        {
            _heartbeatTimer.Stop();
            _watcher.Dispose();
            _trayIcon.Visible = false;
            Application.Exit();
        }
    }
}
