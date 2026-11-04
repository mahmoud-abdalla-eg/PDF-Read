using System;
using System.IO;

namespace PdfRead.Models;

public class RecentFileItem
{
    public string FilePath { get; set; } = string.Empty;
    public string FileName => string.IsNullOrEmpty(FilePath) ? "Untitled" : Path.GetFileName(FilePath);
    public string DirectoryPath => string.IsNullOrEmpty(FilePath) ? string.Empty : Path.GetDirectoryName(FilePath) ?? string.Empty;
    public DateTime LastOpened { get; set; } = DateTime.Now;

    public string LastOpenedFormatted
    {
        get
        {
            var diff = DateTime.Now - LastOpened;
            if (diff.TotalMinutes < 1) return "Just now";
            if (diff.TotalMinutes < 60) return $"{(int)diff.TotalMinutes}m ago";
            if (diff.TotalHours < 24) return $"{(int)diff.TotalHours}h ago";
            if (diff.TotalDays < 7) return $"{(int)diff.TotalDays}d ago";
            return LastOpened.ToString("MMM d, yyyy");
        }
    }

    public string FileSizeFormatted
    {
        get
        {
            try
            {
                if (File.Exists(FilePath))
                {
                    long bytes = new FileInfo(FilePath).Length;
                    if (bytes < 1024) return $"{bytes} B";
                    if (bytes < 1024 * 1024) return $"{bytes / 1024.0:0.#} KB";
                    return $"{bytes / (1024.0 * 1024.0):0.#} MB";
                }
            }
            catch { }
            return "PDF";
        }
    }
}
