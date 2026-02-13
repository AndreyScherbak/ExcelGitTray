using System.Windows;

namespace ExcelGitTray;

public partial class CommitWindow : Window
{
    private readonly GitService _gitService;
    private bool _isBusy;

    public CommitWindow(GitService gitService, string watchedFileName)
    {
        _gitService = gitService;
        InitializeComponent();

        FileLabel.Text = $"{watchedFileName} changed. Enter commit message:";
        Topmost = true;
        Loaded += (_, _) =>
        {
            Activate();
            CommitInput.Focus();
        };
    }

    private async void OnCommit(object sender, RoutedEventArgs e)
    {
        await CommitInternalAsync(pushAfterCommit: false);
    }

    private async void OnCommitPush(object sender, RoutedEventArgs e)
    {
        await CommitInternalAsync(pushAfterCommit: true);
    }

    private async Task CommitInternalAsync(bool pushAfterCommit)
    {
        if (_isBusy)
        {
            return;
        }

        var message = CommitInput.Text.Trim();
        if (string.IsNullOrWhiteSpace(message))
        {
            System.Windows.MessageBox.Show(
                "Commit message cannot be empty.",
                "Validation",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Warning);
            CommitInput.Focus();
            return;
        }

        SetBusy(true);

        try
        {
            var result = await _gitService.CommitAsync(message, pushAfterCommit);
            if (result.Success)
            {
                Close();
                return;
            }

            System.Windows.MessageBox.Show(
                result.Message,
                "Git Operation Failed",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                ex.Message,
                "Unexpected Error",
                System.Windows.MessageBoxButton.OK,
                System.Windows.MessageBoxImage.Error);
        }
        finally
        {
            SetBusy(false);
        }
    }

    private void SetBusy(bool isBusy)
    {
        _isBusy = isBusy;
        CommitButton.IsEnabled = !isBusy;
        CommitPushButton.IsEnabled = !isBusy;
        CommitInput.IsEnabled = !isBusy;
        Cursor = isBusy ? System.Windows.Input.Cursors.Wait : System.Windows.Input.Cursors.Arrow;
    }
}
