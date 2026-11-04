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
