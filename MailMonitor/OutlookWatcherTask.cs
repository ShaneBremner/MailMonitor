using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailMonitor
{
    public class OutlookWatcherTask : IDisposable
    {
        private Outlook.Application? _outlook;
        private Outlook.NameSpace? _session;
        private Outlook.MAPIFolder? _inbox;

        // Held for lifetime of watcher intentionally -- releasing it would unhook events.
        private Outlook.Items? _inboxItems;

        private bool _disposed;

        public void Start()
        {
            try
            {
                _outlook = new Outlook.Application();
                _session = _outlook.Session;
                _inbox = _session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                _inboxItems = _inbox.Items;
                _inboxItems.ItemAdd += OnItemAdded;

                Debug.WriteLine("OutlookWatcherTask started.");
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OutlookWatcherTask failed to start: {ex.Message}");
                Cleanup();
            }
        }

        private void OnItemAdded(object item)
        {
            Outlook.MAPIFolder? destination = null;

            try
            {
                switch (item)
                {
                    case Outlook.MeetingItem meeting:
                        destination = EnsureFolder(_inbox!, "!Process", "_MeetingRequest");
                        meeting.Move(destination);
                        break;

                    case Outlook.MailItem mail:
                        string senderName = SanitizeFolderName(mail.SenderName ?? "Unknown");
                        destination = EnsureFolder(_inbox!, "!Process", "_MessageBySender", senderName);
                        mail.Move(destination);
                        break;

                    // All other item types (appointments, task requests, etc.) are ignored.
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OutlookWatcherTask.OnItemAdded error: {ex.Message}");
            }
            finally
            {
                if (destination != null) Marshal.ReleaseComObject(destination);
                if (item != null) Marshal.ReleaseComObject(item);
            }
        }

        /// <summary>
        /// Walks the folder path from a given root, creating any folders that don't exist.
        /// Returns the final (deepest) folder. Caller is responsible for releasing it.
        /// </summary>
        private Outlook.MAPIFolder EnsureFolder(Outlook.MAPIFolder root, params string[] path)
        {
            // 'current' starts as root but we never release root -- caller owns it.
            Outlook.MAPIFolder current = root;
            Outlook.MAPIFolder? next = null;

            try
            {
                foreach (string segment in path)
                {
                    Outlook.Folders? subFolders = null;
                    try
                    {
                        subFolders = current.Folders;
                        next = FindFolder(subFolders, segment) ?? subFolders.Add(segment);
                    }
                    finally
                    {
                        if (subFolders != null) Marshal.ReleaseComObject(subFolders);
                    }

                    if (current != root) Marshal.ReleaseComObject(current);
                    current = next;
                    next = null;
                }

                return current;
            }
            catch
            {
                if (current != null && current != root) Marshal.ReleaseComObject(current);
                if (next != null) Marshal.ReleaseComObject(next);
                throw;
            }
        }

        /// <summary>
        /// Finds a subfolder by name (case-insensitive). Returns null if not found.
        /// Caller is responsible for releasing the returned folder.
        /// </summary>
        private Outlook.MAPIFolder? FindFolder(Outlook.Folders folders, string name)
        {
            for (int i = 1; i <= folders.Count; i++)
            {
                Outlook.MAPIFolder? candidate = null;
                try
                {
                    candidate = folders[i];
                    if (string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase))
                        return candidate; // Caller releases this.
                }
                catch
                {
                    if (candidate != null) Marshal.ReleaseComObject(candidate);
                    throw;
                }

                // Not a match -- release and keep looking.
                Marshal.ReleaseComObject(candidate);
                candidate = null;
            }

            return null;
        }

        /// <summary>
        /// Strips characters that are invalid in Outlook folder names.
        /// </summary>
        private static string SanitizeFolderName(string name)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
                name = name.Replace(c, '_');

            return name.Trim().TrimStart('~').Replace("  ", " ");
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            Cleanup();
        }

        private void Cleanup()
        {
            if (_inboxItems != null)
            {
                try { _inboxItems.ItemAdd -= OnItemAdded; } catch { }
                Marshal.ReleaseComObject(_inboxItems);
                _inboxItems = null;
            }

            if (_inbox != null) { Marshal.ReleaseComObject(_inbox); _inbox = null; }
            if (_session != null) { Marshal.ReleaseComObject(_session); _session = null; }
            if (_outlook != null) { Marshal.ReleaseComObject(_outlook); _outlook = null; }

            Debug.WriteLine("OutlookWatcherTask cleaned up.");
        }
    }
}
