using System;
using System.IO;
using System.Text.Json;

namespace PdfRead.Models;

public class AppSettings
{
    public string Theme { get; set; } = "dark"; // "dark", "light", "zinc"
    public double DefaultZoom { get; set; } = 1.0;
    public bool ContinuousScroll { get; set; } = true;
    public bool ShowThumbnailsSidebar { get; set; } = true;
    public string DefaultExportFormat { get; set; } = "docx";
    public System.Collections.Generic.List<RecentFileItem> RecentFiles { get; set; } = new();

    public void AddRecentFile(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return;

        RecentFiles.RemoveAll(r => string.Equals(r.FilePath, path, StringComparison.OrdinalIgnoreCase));
        RecentFiles.Insert(0, new RecentFileItem
        {
            FilePath = path,
            LastOpened = DateTime.Now
        });

        if (RecentFiles.Count > 30)
        {
            RecentFiles = RecentFiles.GetRange(0, 30);
        }

        Save();
    }

    public void RemoveRecentFile(string path)
    {
        RecentFiles.RemoveAll(r => string.Equals(r.FilePath, path, StringComparison.OrdinalIgnoreCase));
        Save();
    }

    private static readonly string SettingsPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "PdfRead",
        "settings.json"
    );

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                string json = File.ReadAllText(SettingsPath);
                return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
            }
        }
        catch { }

        return new AppSettings();
    }

    public void Save()
    {
        try
        {
            string dir = Path.GetDirectoryName(SettingsPath)!;
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            string json = JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(SettingsPath, json);
        }
        catch { }
    }
}
