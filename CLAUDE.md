# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Build
dotnet build

# Run (development)
dotnet run --project MailMonitor/MailMonitor.csproj

# Publish self-contained single executable
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
# Output: MailMonitor/bin/Release/net6.0-windows/win-x64/publish/
```

There is no test framework configured in this project.

## Architecture

MailMonitor is a Windows system-tray application (.NET 6 WinForms, C#) that monitors an Outlook inbox in real-time and sends Windows toast notifications for unread mail.

**Main components** (all under `MailMonitor/`):

- **`Program.cs`** — Entry point; initializes WinForms and creates `SilentApp`
- **`SilentApp.cs`** — Extends `ApplicationContext`; owns the tray icon, context menu, heartbeat timer, and starts/stops `OutlookWatcherTask`
- **`SettingsService.cs`** — Static class; persists settings to `HKEY_CURRENT_USER\SOFTWARE\MailMonitor` (currently only `HeartbeatMinutes`, default 5)
- **`OutlookTask.cs`** — Heartbeat: traverses the full Outlook folder tree, counts unread messages, fires a toast notification if any exist
- **`OutlookWatcherTask.cs`** — Real-time: subscribes to `Outlook.Items.ItemAdd`; routes new items into auto-created subfolders (`Inbox → !Process → _MeetingRequest` for meeting requests, `Inbox → !Process → _MessageBySender → {SenderName}` for emails)

**Key patterns:**
- All Outlook COM objects must be released via `Marshal.ReleaseComObject()` — this pattern is pervasive and critical to avoid memory leaks/orphaned Outlook processes.
- The app has no visible window; it lives entirely in the system tray.
- `OutlookWatcherTask` must unsubscribe from COM events before releasing objects (see dispose logic).

**Dependencies:**
- `Microsoft.Office.Interop.Outlook` v9.4 — COM interop; `EmbedInteropTypes=true`
- `Microsoft.Toolkit.Uwp.Notifications` v7.1.3 — Toast notifications
- `Microsoft.VisualBasic` — `InputBox` for the interval-setting dialog
