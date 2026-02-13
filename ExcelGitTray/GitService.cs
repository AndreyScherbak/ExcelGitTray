using System.Diagnostics;
using System.IO;
using System.Text;

namespace ExcelGitTray;

public sealed class GitService
{
    private readonly string _repoPath;
    private readonly string _targetRelativePath;
    private readonly SemaphoreSlim _gitLock = new(1, 1);

    public GitService(string repoPath, string targetFilePath)
    {
        if (string.IsNullOrWhiteSpace(repoPath))
        {
            throw new ArgumentException("Repository path cannot be empty.", nameof(repoPath));
        }

        if (!Directory.Exists(repoPath))
        {
            throw new DirectoryNotFoundException($"Repository directory not found: {repoPath}");
        }

        _repoPath = repoPath;
        _targetRelativePath = Path.GetRelativePath(repoPath, targetFilePath);
    }

    public async Task<GitResult> PullAsync(CancellationToken cancellationToken = default)
    {
        await _gitLock.WaitAsync(cancellationToken);

        try
        {
            var hasPendingChanges = await HasPendingChangesInternalAsync(cancellationToken);
            if (hasPendingChanges)
            {
                return GitResult.FailureResult(
                    "Pull blocked: local or staged changes exist for the Excel file. Commit/stash/discard them first.");
            }

            var result = await RunGitAsync(new[] { "pull", "--ff-only" }, cancellationToken);
            return result.Success
                ? GitResult.SuccessResult("Repository updated successfully.")
                : GitResult.FailureResult($"git pull failed: {result.Summary}");
        }
        finally
        {
            _gitLock.Release();
        }
    }

    public async Task<GitResult> CommitAsync(string message, bool pushAfterCommit, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(message))
        {
            return GitResult.FailureResult("Commit message cannot be empty.");
        }

        await _gitLock.WaitAsync(cancellationToken);

        try
        {
            var addResult = await RunGitAsync(new[] { "add", "--", _targetRelativePath }, cancellationToken);
            if (!addResult.Success)
            {
                return GitResult.FailureResult($"git add failed: {addResult.Summary}");
            }

            var diffResult = await RunGitAsync(
                new[] { "diff", "--cached", "--quiet" },
                cancellationToken,
                treatNonZeroExitCodeAsError: false);

            if (diffResult.ExitCode == 0)
            {
                return GitResult.FailureResult("No staged changes detected. Nothing to commit.");
            }

            if (diffResult.ExitCode != 1)
            {
                return GitResult.FailureResult($"Unable to check staged changes: {diffResult.Summary}");
            }

            var commitResult = await RunGitAsync(new[] { "commit", "-m", message.Trim() }, cancellationToken);
            if (!commitResult.Success)
            {
                return GitResult.FailureResult($"git commit failed: {commitResult.Summary}");
            }

            if (!pushAfterCommit)
            {
                return GitResult.SuccessResult("Commit completed.");
            }

            var pushResult = await RunGitAsync(new[] { "push" }, cancellationToken);
            return pushResult.Success
                ? GitResult.SuccessResult("Commit and push completed.")
                : GitResult.FailureResult($"Commit succeeded, but git push failed: {pushResult.Summary}");
        }
        finally
        {
            _gitLock.Release();
        }
    }

    public async Task<GitResult> PushAsync(CancellationToken cancellationToken = default)
    {
        await _gitLock.WaitAsync(cancellationToken);

        try
        {
            var result = await RunGitAsync(new[] { "push" }, cancellationToken);
            return result.Success
                ? GitResult.SuccessResult("Push completed.")
                : GitResult.FailureResult($"git push failed: {result.Summary}");
        }
        finally
        {
            _gitLock.Release();
        }
    }

    public async Task<bool> HasPendingChangesAsync(CancellationToken cancellationToken = default)
    {
        await _gitLock.WaitAsync(cancellationToken);

        try
        {
            return await HasPendingChangesInternalAsync(cancellationToken);
        }
        finally
        {
            _gitLock.Release();
        }
    }

    private async Task<ProcessResult> RunGitAsync(
        IReadOnlyList<string> arguments,
        CancellationToken cancellationToken,
        bool treatNonZeroExitCodeAsError = true)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = "git",
            WorkingDirectory = _repoPath,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        foreach (var arg in arguments)
        {
            startInfo.ArgumentList.Add(arg);
        }

        using var process = new Process { StartInfo = startInfo };

        try
        {
            process.Start();
        }
        catch (Exception ex)
        {
            return new ProcessResult(-1, false, string.Empty, ex.Message);
        }

        var standardOutputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        var standardErrorTask = process.StandardError.ReadToEndAsync(cancellationToken);

        await process.WaitForExitAsync(cancellationToken);

        var output = await standardOutputTask;
        var error = await standardErrorTask;
        var success = process.ExitCode == 0 || !treatNonZeroExitCodeAsError;

        return new ProcessResult(process.ExitCode, success, output, error);
    }

    private async Task<bool> HasPendingChangesInternalAsync(CancellationToken cancellationToken)
    {
        var statusResult = await RunGitAsync(
            new[] { "status", "--porcelain", "--", _targetRelativePath },
            cancellationToken);

        if (!statusResult.Success)
        {
            throw new InvalidOperationException($"Unable to read git status: {statusResult.Summary}");
        }

        return !string.IsNullOrWhiteSpace(statusResult.StandardOutput);
    }

    private readonly record struct ProcessResult(int ExitCode, bool Success, string StandardOutput, string StandardError)
    {
        public string Summary
        {
            get
            {
                var builder = new StringBuilder();

                if (!string.IsNullOrWhiteSpace(StandardError))
                {
                    builder.Append(StandardError.Trim());
                }

                if (!string.IsNullOrWhiteSpace(StandardOutput))
                {
                    if (builder.Length > 0)
                    {
                        builder.Append(" | ");
                    }

                    builder.Append(StandardOutput.Trim());
                }

                if (builder.Length == 0)
                {
                    builder.Append($"Exit code: {ExitCode}");
                }

                return builder.ToString();
            }
        }
    }
}

public sealed record GitResult(bool Success, string Message)
{
    public static GitResult SuccessResult(string message) => new(true, message);

    public static GitResult FailureResult(string message) => new(false, message);
}
