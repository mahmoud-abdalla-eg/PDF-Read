using System;
using System.Windows;
using PdfRead.Services;

namespace PdfRead;

public partial class SettingsWindow : Window
{
    public SettingsWindow()
    {
        InitializeComponent();
        RefreshStatus();
    }

    private void RefreshStatus()
    {
        bool pdfMenu = ShellIntegration.IsPdfContextMenuRegistered();
        TxtPdfMenuStatus.Text = pdfMenu ? "✓ Active" : "Disabled";
        BtnTogglePdfMenu.Content = pdfMenu ? "Remove" : "Enable";

        bool officeMenu = ShellIntegration.IsOfficeConversionContextMenuRegistered();
        TxtOfficeMenuStatus.Text = officeMenu ? "✓ Active" : "Disabled";
        BtnToggleOfficeMenu.Content = officeMenu ? "Remove" : "Enable";

        bool win11 = ShellIntegration.IsClassicContextMenuEnabled();
        TxtWin11Status.Text = win11 ? "✓ Enabled" : "Default";
        BtnToggleWin11.Content = win11 ? "Revert" : "Enable Direct";

        bool startup = ShellIntegration.IsStartupEnabled();
        TxtStartupStatus.Text = startup ? "✓ Enabled" : "Disabled";
        BtnToggleStartup.Content = startup ? "Disable" : "Enable";
    }

    private void ShowBanner(string message)
    {
        TxtStatusBanner.Text = message;
        StatusBanner.Visibility = Visibility.Visible;
    }

    private void BtnUpdateDesktopShortcut_Click(object sender, RoutedEventArgs e)
    {
        ShellIntegration.CreateDesktopShortcut();
        ShowBanner("✓ Desktop shortcut updated with red PDF badge icon.");
    }

    private void BtnUpdateStartMenuShortcut_Click(object sender, RoutedEventArgs e)
    {
        ShellIntegration.CreateStartMenuShortcut();
        ShowBanner("✓ Start Menu & Taskbar shortcut updated with red PDF badge icon.");
    }

    private void BtnToggleStartup_Click(object sender, RoutedEventArgs e)
    {
        bool current = ShellIntegration.IsStartupEnabled();
        ShellIntegration.SetStartup(!current);
        RefreshStatus();
        ShowBanner(current ? "Startup launch disabled." : "✓ Launch on Windows Startup enabled.");
    }

    private void BtnUpdateAllShortcuts_Click(object sender, RoutedEventArgs e)
    {
        ShellIntegration.UpdateAllIconsAndShortcuts();
        ShowBanner("✓ All shortcuts (Desktop, Start Menu, Taskbar) updated with red PDF badge icon.");
    }

    private void BtnTogglePdfMenu_Click(object sender, RoutedEventArgs e)
    {
        bool current = ShellIntegration.IsPdfContextMenuRegistered();
        ShellIntegration.SetPdfContextMenu(!current);
        RefreshStatus();
        ShowBanner(current ? "PDF context menu removed." : "✓ 'Open with PdfRead' context menu registered.");
    }

    private void BtnToggleOfficeMenu_Click(object sender, RoutedEventArgs e)
    {
        bool current = ShellIntegration.IsOfficeConversionContextMenuRegistered();
        ShellIntegration.SetOfficeConversionContextMenu(!current);
        RefreshStatus();
        ShowBanner(current ? "Office conversion context menu removed." : "✓ 'Convert to PDF with PdfRead' context menu registered.");
    }

    private void BtnToggleWin11_Click(object sender, RoutedEventArgs e)
    {
        bool current = ShellIntegration.IsClassicContextMenuEnabled();
        ShellIntegration.SetClassicContextMenu(!current);
        RefreshStatus();
        ShowBanner(current ? "Reverted to default Windows 11 menu." : "✓ Windows 11 direct first-click menu enabled.");
    }

    private void BtnSetDefault_Click(object sender, RoutedEventArgs e)
    {
        ShellIntegration.PromptSetAsDefault();
        ShowBanner("Windows Default Apps dialog opened.");
    }

    private void BtnRemoveAll_Click(object sender, RoutedEventArgs e)
    {
        if (MessageBox.Show("Are you sure you want to remove all PdfRead registry keys, context menus, and associations?", 
            "Confirm Removal", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
        {
            ShellIntegration.UnregisterFileAssociations();
            RefreshStatus();
            ShowBanner("All PdfRead context menus and registry entries removed cleanly.");
        }
    }

    private void BtnClose_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
