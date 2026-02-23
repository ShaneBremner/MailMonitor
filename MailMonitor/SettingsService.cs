using Microsoft.Win32;
using System;

namespace MailMonitor
{
    public static class SettingsService
    {
        private const string RegistryKeyPath = @"SOFTWARE\MailMonitor";

        private const int DefaultHeartbeatMinutes = 5;

        public static int HeartbeatMinutes
        {
            get => GetIntValue(nameof(HeartbeatMinutes), DefaultHeartbeatMinutes);
            set => SetIntValue(nameof(HeartbeatMinutes), value);
        }

        private static int GetIntValue(string name, int defaultValue)
        {
            using var key = Registry.CurrentUser.OpenSubKey(RegistryKeyPath);
            if (key?.GetValue(name) is int val)
                return val;
            return defaultValue;
        }

        private static void SetIntValue(string name, int value)
        {
            using var key = Registry.CurrentUser.CreateSubKey(RegistryKeyPath, writable: true);
            key.SetValue(name, value, RegistryValueKind.DWord);
        }
    }
}
