using System.IO;
using System.Text.Json;

namespace ExcelGitTray;

public sealed class AppConfig
{
    public string ExcelFilePath { get; set; } = @"C:\Repo\schedule.xlsx";
    public string TrayIconPath { get; set; } = "tray.ico";
    private static readonly string ConfigPath = Path.Combine(AppContext.BaseDirectory, "appsettings.json");

    public static AppConfig LoadOrCreateDefault()
    {
        if (!File.Exists(ConfigPath))
        {
            var defaultConfig = new AppConfig();
            defaultConfig.Save();
            return defaultConfig;
        }

        var json = File.ReadAllText(ConfigPath);
        var config = JsonSerializer.Deserialize<AppConfig>(json) ?? new AppConfig();

        if (string.IsNullOrWhiteSpace(config.ExcelFilePath))
        {
            throw new InvalidOperationException("ExcelFilePath is missing in appsettings.json.");
        }

        return config;
    }

    public void Save()
    {
        var json = JsonSerializer.Serialize(this, new JsonSerializerOptions
        {
            WriteIndented = true
        });

        File.WriteAllText(ConfigPath, json);
    }
}
