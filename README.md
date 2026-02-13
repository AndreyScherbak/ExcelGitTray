# ExcelGitTray

ExcelGitTray is a Windows tray application that watches an Excel file in a Git repository and helps you commit, push, and pull changes directly from the system tray.

## Features

- Watches a selected Excel file for changes.
- Debounces file events and waits for file unlock before prompting.
- Opens a commit window when pending changes are detected.
- Supports `Commit`, `Commit + Push`, and tray `Pull` operations.
- Blocks pull while local Excel changes exist or the file is in use.
- Lets you open the watched file and switch to a different Excel file from tray menu.

## Tech Stack

- .NET `net10.0-windows`
- WPF + Windows Forms `NotifyIcon`
- Git CLI (`git`) for repository operations

## Prerequisites

- Windows
- .NET SDK 10.0+
- Git installed and available in `PATH`
- An Excel workbook located inside a Git working tree

## Configuration

Create `ExcelGitTray/appsettings.json` with your local paths:

```json
{
  "ExcelFilePath": "C:\\Path\\To\\Your\\Workbook.xlsm",
  "TrayIconPath": "tray.ico"
}
```

Notes:
- `ExcelFilePath` should point to the workbook you want to watch.
- `TrayIconPath` can be absolute or relative to the app output folder.

## Run Locally

```powershell
dotnet restore
dotnet run --project .\ExcelGitTray\ExcelGitTray.csproj
```

Or build the solution:

```powershell
dotnet build .\ExcelGitTray.sln
```

## Tray Menu

- `Open`: Opens the configured Excel file.
- `Pull`: Runs `git pull --ff-only` for the workbook repo.
- `Pull + Open`: Pulls first, then opens the file on success.
- `Set Excel File...`: Select a different workbook to watch.
- `Exit`: Stops watcher and closes the tray app.

## Project Structure

```text
ExcelGitTray.sln
ExcelGitTray/
  App.xaml.cs           # Startup, tray, watcher lifecycle
  AppConfig.cs          # appsettings.json load/save
  GitService.cs         # git pull/commit/push/status helpers
  CommitWindow.xaml.cs  # commit prompt window logic
  appsettings.json      # local machine config (ignored by git)
```

## Git Notes

This repository includes a `.gitignore` configured for:
- Visual Studio and IDE metadata
- .NET build artifacts (`bin/`, `obj/`)
- local `ExcelGitTray/appsettings.json`

If build artifacts are already tracked, untrack them once:

```powershell
git rm -r --cached .\ExcelGitTray\bin .\ExcelGitTray\obj
git commit -m "chore: stop tracking build artifacts"
```
