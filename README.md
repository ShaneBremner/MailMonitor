# MailMonitor

A silent Windows system-tray app that monitors your Outlook inbox.

## Features

- **Heartbeat task** — periodically scans your inbox and subfolders and shows a Windows toast notification with the total unread message count.
- **Inbox watcher** — listens for new items in real time:
  - New **meeting requests** are moved to `Inbox → !Process → _MeetingRequest`
  - New **emails** are moved to `Inbox → !Process → _MessageBySender → {SenderName}`
- **Configurable interval** — right-click the tray icon → *Set Interval* to change the heartbeat period. The setting is persisted to the registry.

## Requirements

- Windows 10 or 11
- .NET 6 SDK — https://dotnet.microsoft.com/download/dotnet/6.0
- Microsoft Outlook installed and configured

## Build & Run

```
cd MailMonitor
dotnet build
dotnet run
```

Or publish a self-contained executable:

```
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true
```

The output will be in `bin\Release\net6.0-windows\win-x64\publish\`.

## Run on Windows Startup

1. Press `Win + R`, type `shell:startup`, press Enter.
2. Create a shortcut to `MailMonitor.exe` in that folder.

## Settings (Registry)

Settings are stored under `HKEY_CURRENT_USER\SOFTWARE\MailMonitor`.

| Value Name         | Type  | Default | Description                        |
|--------------------|-------|---------|------------------------------------|
| HeartbeatMinutes   | DWORD | 5       | Interval between unread mail checks |

## Project Structure

| File                    | Purpose                                      |
|-------------------------|----------------------------------------------|
| `Program.cs`            | Entry point                                  |
| `SilentApp.cs`          | ApplicationContext, tray icon, timer         |
| `SettingsService.cs`    | Registry-backed settings                     |
| `OutlookTask.cs`        | Heartbeat: counts and notifies unread mail   |
| `OutlookWatcherTask.cs` | Startup: watches inbox and sorts new items   |
