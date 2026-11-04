using System;
using System.Diagnostics;
using System.IO;
using Microsoft.Win32;

namespace PdfRead.Services;

public static class ShellIntegration
{
    private const string ProgId = "PdfRead.Document";
    private const string AppName = "PdfRead.exe";

    public static string GetExecutablePath()
    {
        return Environment.ProcessPath 
               ?? Process.GetCurrentProcess().MainModule?.FileName 
               ?? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, AppName);
    }

    public static bool IsPdfContextMenuRegistered()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Classes\SystemFileAssociations\.pdf\shell\Open with PdfRead");
            return key != null;
        }
        catch { return false; }
    }

    public static bool IsOfficeConversionContextMenuRegistered()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Classes\SystemFileAssociations\.docx\shell\Convert to PDF with PdfRead");
            return key != null;
        }
        catch { return false; }
    }

    public static void SetPdfContextMenu(bool enable)
    {
        string exePath = GetExecutablePath();

        if (enable)
        {
            using var ctxKey = Registry.CurrentUser.CreateSubKey(@"Software\Classes\SystemFileAssociations\.pdf\shell\Open with PdfRead");
            ctxKey.SetValue("", "Open with PdfRead");
            ctxKey.SetValue("Icon", $"\"{exePath}\",0");
            using var cmdKey = ctxKey.CreateSubKey("command");
            cmdKey.SetValue("", $"\"{exePath}\" \"%1\"");
        }
        else
        {
            try { Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\SystemFileAssociations\.pdf\shell\Open with PdfRead", false); } catch { }
        }
    }

    public static void SetOfficeConversionContextMenu(bool enable)
    {
        string exePath = GetExecutablePath();
        string[] exts = { ".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt" };

        foreach (var ext in exts)
        {
            string subKeyPath = $@"Software\Classes\SystemFileAssociations\{ext}\shell\Convert to PDF with PdfRead";
            if (enable)
            {
                using var ctxKey = Registry.CurrentUser.CreateSubKey(subKeyPath);
                ctxKey.SetValue("", "Convert to PDF with PdfRead");
                ctxKey.SetValue("Icon", $"\"{exePath}\",0");
                using var cmdKey = ctxKey.CreateSubKey("command");
                cmdKey.SetValue("", $"\"{exePath}\" --convert-to-pdf \"%1\"");
            }
            else
            {
                try { Registry.CurrentUser.DeleteSubKeyTree(subKeyPath, false); } catch { }
            }
        }
    }

    public static void RegisterFileAssociations()
    {
        string exePath = GetExecutablePath();

        using (var progKey = Registry.CurrentUser.CreateSubKey($@"Software\Classes\{ProgId}"))
        {
            progKey.SetValue("", "PDF Document (PdfRead)");
            using var iconKey = progKey.CreateSubKey("DefaultIcon");
            iconKey.SetValue("", $"\"{exePath}\",0");

            using var cmdKey = progKey.CreateSubKey(@"shell\open\command");
            cmdKey.SetValue("", $"\"{exePath}\" \"%1\"");
        }

        SetPdfContextMenu(true);
        SetOfficeConversionContextMenu(true);
        UpdateAllIconsAndShortcuts();
    }

    public static void UnregisterFileAssociations()
    {
        try { Registry.CurrentUser.DeleteSubKeyTree($@"Software\Classes\{ProgId}", false); } catch { }
        SetPdfContextMenu(false);
        SetOfficeConversionContextMenu(false);
    }

    public static void CreateShortcut(string targetLnkPath, string? exePath = null, string? description = null)
    {
        try
        {
            exePath ??= GetExecutablePath();
            Type? shellType = Type.GetTypeFromProgID("WScript.Shell");
            if (shellType == null) return;

            dynamic? shell = Activator.CreateInstance(shellType);
            if (shell == null) return;

            string? dir = Path.GetDirectoryName(targetLnkPath);
            if (!string.IsNullOrEmpty(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            dynamic shortcut = shell.CreateShortcut(targetLnkPath);
            shortcut.TargetPath = exePath;
            shortcut.WorkingDirectory = Path.GetDirectoryName(exePath);
            shortcut.Description = description ?? "PdfRead - High-Performance PDF Reader & Editor";
            shortcut.IconLocation = $"{exePath},0";
            shortcut.Save();
        }
        catch { }
    }

    public static void CreateDesktopShortcut()
    {
        string desktopDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        CreateShortcut(Path.Combine(desktopDir, "PdfRead.lnk"));
    }

    public static void CreateStartMenuShortcut()
    {
        string programsDir = Environment.GetFolderPath(Environment.SpecialFolder.Programs);
        CreateShortcut(Path.Combine(programsDir, "PdfRead.lnk"));
    }

    public static void UpdateAllIconsAndShortcuts()
    {
        CreateDesktopShortcut();
        CreateStartMenuShortcut();
    }

    private const string StartupRunKey = @"Software\Microsoft\Windows\CurrentVersion\Run";

    public static bool IsStartupEnabled()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(StartupRunKey);
            return key?.GetValue("PdfRead") != null;
        }
        catch { return false; }
    }

    public static void SetStartup(bool enable)
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(StartupRunKey, true);
            if (key == null) return;

            if (enable)
            {
                key.SetValue("PdfRead", $"\"{GetExecutablePath()}\"");
            }
            else
            {
                key.DeleteValue("PdfRead", false);
            }
        }
        catch { }
    }

    public static bool IsClassicContextMenuEnabled()
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(@"Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32");
            return key != null;
        }
        catch { return false; }
    }

    public static void SetClassicContextMenu(bool enable)
    {
        try
        {
            if (enable)
            {
                using var key = Registry.CurrentUser.CreateSubKey(@"Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32");
                key.SetValue("", "");
            }
            else
            {
                Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}", false);
            }
        }
        catch { }
    }

    public static void PromptSetAsDefault()
    {
        try
        {
            RegisterFileAssociations();
            string sampleFile = Path.Combine(Path.GetTempPath(), "sample_document.pdf");
            if (!File.Exists(sampleFile))
            {
                File.WriteAllBytes(sampleFile, new byte[100]);
            }

            Process.Start(new ProcessStartInfo
            {
                FileName = "rundll32.exe",
                Arguments = $"shell32.dll,OpenAs_RunDLL \"{sampleFile}\"",
                UseShellExecute = true
            });
        }
        catch { }
    }
}
