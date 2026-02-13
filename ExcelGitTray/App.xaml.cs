using System.IO;
using System.Threading;
using System.Windows;
using System.Windows.Threading;
using Forms = System.Windows.Forms;

namespace ExcelGitTray;

public partial class App : System.Windows.Application
{
    private static readonly TimeSpan DebounceDelay = TimeSpan.FromMilliseconds(900);
    private static readonly TimeSpan UnlockTimeout = TimeSpan.FromSeconds(25);

    private readonly object _debounceLock = new();

    private AppConfig? _config;
    private Forms.NotifyIcon? _trayIcon;
    private Forms.ContextMenuStrip? _trayMenu;
    private FileSystemWatcher? _watcher;
    private CancellationTokenSource? _debounceCts;
    private GitService? _gitService;
    private System.Drawing.Icon? _trayIconResource;
    private CommitWindow? _commitWindow;
    private int _suppressCommitPrompts;

    private void Application_Startup(object sender, StartupEventArgs e)
    {
        try
        {
            _config = AppConfig.LoadOrCreateDefault();
            ConfigureTargetFilePath(_config.ExcelFilePath, saveConfig: false);
            CreateTrayIcon();
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                ex.Message,
                "ExcelGitTray Startup Error",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);
            Shutdown();
        }
    }

    private void CreateTrayIcon()
    {
        _trayMenu = new Forms.ContextMenuStrip();
        _trayMenu.Items.Add("Open", null, (_, _) => OpenExcelFile());
        _trayMenu.Items.Add("Pull", null, async (_, _) => await PullAsync());
        _trayMenu.Items.Add("Pull + Open", null, async (_, _) => await PullAndOpenAsync());
        _trayMenu.Items.Add("Set Excel File...", null, (_, _) => ConfigureExcelFileFromDialog());
        _trayMenu.Items.Add("-");
        _trayMenu.Items.Add("Exit", null, (_, _) => ExitApplication());

        _trayIconResource = TryLoadTrayIcon(_config?.TrayIconPath);

        _trayIcon = new Forms.NotifyIcon
        {
            Icon = _trayIconResource ?? System.Drawing.SystemIcons.Application,
            Text = "Excel Git Tray",
            ContextMenuStrip = _trayMenu,
            Visible = true
        };

        _trayIcon.DoubleClick += (_, _) => BringCommitWindowToFront();
    }

    private System.Drawing.Icon? TryLoadTrayIcon(string? configuredPath)
    {
        if (string.IsNullOrWhiteSpace(configuredPath))
        {
            return null;
        }

        var resolvedPath = ResolvePath(configuredPath);
        if (!File.Exists(resolvedPath))
        {
            ShowTrayBalloon(
                "Tray icon not found",
                $"Icon file does not exist: {resolvedPath}",
                Forms.ToolTipIcon.Warning);
            return null;
        }

        try
        {
            return new System.Drawing.Icon(resolvedPath);
        }
        catch (Exception ex)
        {
            ShowTrayBalloon(
                "Tray icon error",
                $"Could not load icon: {ex.Message}",
                Forms.ToolTipIcon.Warning);
            return null;
        }
    }

    private void StartWatcher(string filePath)
    {
        StopWatcher();

        if (!File.Exists(filePath))
        {
            ShowTrayBalloon(
                "Excel file not found",
                $"Watcher started, but file is missing: {filePath}",
                Forms.ToolTipIcon.Warning);
        }

        var directory = Path.GetDirectoryName(filePath);
        var fileName = Path.GetFileName(filePath);

        if (string.IsNullOrWhiteSpace(directory) || string.IsNullOrWhiteSpace(fileName))
        {
            throw new InvalidOperationException($"Invalid watch target: {filePath}");
        }

        _watcher = new FileSystemWatcher(directory, fileName)
        {
            NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.Size | NotifyFilters.FileName,
            EnableRaisingEvents = true,
            IncludeSubdirectories = false
        };

        _watcher.Changed += OnFileChanged;
        _watcher.Created += OnFileChanged;
        _watcher.Renamed += OnFileChanged;
    }

    private void StopWatcher()
    {
        if (_watcher is null)
        {
            return;
        }

        _watcher.EnableRaisingEvents = false;
        _watcher.Changed -= OnFileChanged;
        _watcher.Created -= OnFileChanged;
        _watcher.Renamed -= OnFileChanged;
        _watcher.Dispose();
        _watcher = null;
    }

    private void OnFileChanged(object sender, FileSystemEventArgs e)
    {
        if (IsCommitPromptSuppressed())
        {
            return;
        }

        lock (_debounceLock)
        {
            _debounceCts?.Cancel();
            _debounceCts?.Dispose();
            _debounceCts = new CancellationTokenSource();
            _ = HandleChangeDebouncedAsync(_debounceCts.Token);
        }
    }

    private async Task HandleChangeDebouncedAsync(CancellationToken cancellationToken)
    {
        try
        {
            if (_config is null)
            {
                return;
            }

            if (IsCommitPromptSuppressed())
            {
                return;
            }

            await Task.Delay(DebounceDelay, cancellationToken);

            if (IsCommitPromptSuppressed())
            {
                return;
            }

            var unlocked = await WaitForFileUnlockAsync(_config.ExcelFilePath, UnlockTimeout, cancellationToken);
            if (!unlocked)
            {
                ShowTrayBalloon(
                    "File still locked",
                    "Excel file remained locked after save. Try again after Excel finishes writing.",
                    Forms.ToolTipIcon.Warning);
                return;
            }

            if (_gitService is null)
            {
                return;
            }

            var hasPendingChanges = await _gitService.HasPendingChangesAsync(cancellationToken);
            if (!hasPendingChanges)
            {
                return;
            }

            await Dispatcher.InvokeAsync(ShowCommitWindow, DispatcherPriority.Normal, cancellationToken);
        }
        catch (OperationCanceledException)
        {
            // Newer file event superseded this one.
        }
        catch (Exception ex)
        {
            ShowTrayBalloon("Watcher error", ex.Message, Forms.ToolTipIcon.Error);
        }
    }

    private static async Task<bool> WaitForFileUnlockAsync(string path, TimeSpan timeout, CancellationToken cancellationToken)
    {
        var deadline = DateTime.UtcNow + timeout;

        while (DateTime.UtcNow < deadline)
        {
            cancellationToken.ThrowIfCancellationRequested();

            try
            {
                using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                return true;
            }
            catch (IOException)
            {
            }
            catch (UnauthorizedAccessException)
            {
            }

            await Task.Delay(250, cancellationToken);
        }

        return false;
    }

    private async Task<GitResult?> PullAsync()
    {
        if (_gitService is null || _config is null)
        {
            return null;
        }

        if (IsFileLocked(_config.ExcelFilePath))
        {
            var lockMessage = "Pull blocked: Excel file is currently in use. Save/close it and try again.";
            ShowTrayBalloon("Pull blocked", lockMessage, Forms.ToolTipIcon.Warning);
            return GitResult.FailureResult(lockMessage);
        }

        using var _ = BeginSuppressCommitPrompt();
        var result = await _gitService.PullAsync();
        await Task.Delay(DebounceDelay);
        ShowTrayBalloon(
            result.Success ? "Pull succeeded" : "Pull failed",
            result.Message,
            result.Success ? Forms.ToolTipIcon.Info : Forms.ToolTipIcon.Warning);
        return result;
    }

    private async Task PullAndOpenAsync()
    {
        var result = await PullAsync();
        if (result?.Success == true)
        {
            OpenExcelFile();
        }
    }

    private void OpenExcelFile()
    {
        if (_config is null)
        {
            return;
        }

        var filePath = _config.ExcelFilePath;
        if (!File.Exists(filePath))
        {
            ShowTrayBalloon("Open failed", $"Excel file not found: {filePath}", Forms.ToolTipIcon.Error);
            return;
        }

        try
        {
            var startInfo = new System.Diagnostics.ProcessStartInfo
            {
                FileName = filePath,
                UseShellExecute = true
            };

            System.Diagnostics.Process.Start(startInfo);
        }
        catch (Exception ex)
        {
            ShowTrayBalloon("Open failed", ex.Message, Forms.ToolTipIcon.Error);
        }
    }

    private void ConfigureExcelFileFromDialog()
    {
        if (_config is null)
        {
            return;
        }

        var dialog = new Microsoft.Win32.OpenFileDialog
        {
            Title = "Select Excel file",
            Filter = "Excel Files (*.xlsx;*.xlsm;*.xlsb;*.xls)|*.xlsx;*.xlsm;*.xlsb;*.xls|All Files (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false,
            FileName = Path.GetFileName(_config.ExcelFilePath),
            InitialDirectory = GetInitialDirectory(_config.ExcelFilePath)
        };

        if (dialog.ShowDialog() != true)
        {
            return;
        }

        try
        {
            using var _ = BeginSuppressCommitPrompt();
            ConfigureTargetFilePath(dialog.FileName, saveConfig: true);
            ShowTrayBalloon("Excel file updated", dialog.FileName, Forms.ToolTipIcon.Info);
        }
        catch (Exception ex)
        {
            ShowTrayBalloon("Update failed", ex.Message, Forms.ToolTipIcon.Error);
        }
    }

    private static string? GetInitialDirectory(string filePath)
    {
        var directory = Path.GetDirectoryName(filePath);
        return Directory.Exists(directory) ? directory : null;
    }

    private void ConfigureTargetFilePath(string targetFilePath, bool saveConfig)
    {
        var repoPath = Path.GetDirectoryName(targetFilePath);
        if (string.IsNullOrWhiteSpace(repoPath))
        {
            throw new InvalidOperationException($"Could not determine repository path from {targetFilePath}");
        }

        if (_config is null)
        {
            throw new InvalidOperationException("Application config is not available.");
        }

        _gitService = new GitService(repoPath, targetFilePath);
        _config.ExcelFilePath = targetFilePath;

        if (saveConfig)
        {
            _config.Save();
        }

        if (_commitWindow is not null)
        {
            _commitWindow.Close();
        }

        StartWatcher(targetFilePath);
    }

    private void ShowCommitWindow()
    {
        if (_gitService is null || _config is null)
        {
            return;
        }

        if (_commitWindow is not null)
        {
            BringCommitWindowToFront();
            return;
        }

        _commitWindow = new CommitWindow(_gitService, Path.GetFileName(_config.ExcelFilePath));
        _commitWindow.Closed += (_, _) => _commitWindow = null;
        _commitWindow.Show();
        BringCommitWindowToFront();
    }

    private void BringCommitWindowToFront()
    {
        if (_commitWindow is null)
        {
            return;
        }

        _commitWindow.Topmost = true;
        _commitWindow.Activate();
        _commitWindow.Focus();
    }

    private void ShowTrayBalloon(string title, string message, Forms.ToolTipIcon icon)
    {
        _trayIcon?.ShowBalloonTip(3000, title, message, icon);
    }

    private void ExitApplication()
    {
        StopWatcher();
        _debounceCts?.Cancel();
        _debounceCts?.Dispose();
        _trayIcon?.Dispose();
        _trayIconResource?.Dispose();
        Shutdown();
    }

    private static string ResolvePath(string configuredPath)
    {
        if (Path.IsPathRooted(configuredPath))
        {
            return configuredPath;
        }

        return Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, configuredPath));
    }

    private bool IsCommitPromptSuppressed() => Volatile.Read(ref _suppressCommitPrompts) > 0;

    private IDisposable BeginSuppressCommitPrompt()
    {
        Interlocked.Increment(ref _suppressCommitPrompts);
        return new ScopeGuard(() => Interlocked.Decrement(ref _suppressCommitPrompts));
    }

    private static bool IsFileLocked(string path)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            return false;
        }
        catch (IOException)
        {
            return true;
        }
        catch (UnauthorizedAccessException)
        {
            return true;
        }
    }

    private sealed class ScopeGuard(Action onDispose) : IDisposable
    {
        private readonly Action _onDispose = onDispose;
        private int _disposed;

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) == 0)
            {
                _onDispose();
            }
        }
    }
}
