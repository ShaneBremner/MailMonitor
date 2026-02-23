using Microsoft.Toolkit.Uwp.Notifications;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MailMonitor
{
    public static class OutlookTask
    {
        public static void CheckUnreadMail()
        {
            Outlook.Application? outlook = null;
            Outlook.NameSpace? session = null;
            Outlook.MAPIFolder? inbox = null;

            try
            {
                outlook = new Outlook.Application();
                session = outlook.Session;
                inbox = session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                int totalUnread = 0;
                ScanFolder(inbox, ref totalUnread);

                if (totalUnread > 0)
                    ShowNotification(totalUnread);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"OutlookTask.CheckUnreadMail error: {ex.Message}");
            }
            finally
            {
                if (inbox != null) Marshal.ReleaseComObject(inbox);
                if (session != null) Marshal.ReleaseComObject(session);
                if (outlook != null) Marshal.ReleaseComObject(outlook);
            }
        }

        private static void ScanFolder(Outlook.MAPIFolder folder, ref int totalUnread)
        {
            Outlook.Folders? subFolders = null;

            try
            {
                totalUnread += folder.UnReadItemCount;

                subFolders = folder.Folders;
                for (int i = 1; i <= subFolders.Count; i++)
                {
                    Outlook.MAPIFolder? sub = null;
                    try
                    {
                        sub = subFolders[i];
                        ScanFolder(sub, ref totalUnread);
                    }
                    finally
                    {
                        if (sub != null) Marshal.ReleaseComObject(sub);
                    }
                }
            }
            finally
            {
                if (subFolders != null) Marshal.ReleaseComObject(subFolders);
            }
        }

        private static void ShowNotification(int totalUnread)
        {
            new ToastContentBuilder()
                .AddText("📬 Unread Mail")
                .AddText($"You have {totalUnread} unread message{(totalUnread == 1 ? "" : "s")} across your inbox and subfolders.")
                .Show();
        }
    }
}
