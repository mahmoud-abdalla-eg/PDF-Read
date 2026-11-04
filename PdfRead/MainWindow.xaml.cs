using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using Microsoft.Web.WebView2.Core;
using Microsoft.Win32;
using PdfRead.Models;
using PdfRead.Services;

namespace PdfRead;

public partial class MainWindow : Window
{
    private readonly ObservableCollection<PdfDocumentTab> _tabs = new();
    private PdfDocumentTab? _activeTab;
    private readonly OfficeConversionService _officeConverter = new();
    private readonly DocumentExportService _exportService = new();
    private readonly PdfEditorService _editorService = new();
    private readonly List<TextBlockItem> _currentTextBlocks = new();
    private bool _isWebViewInitialized;

    public MainWindow()
    {
        InitializeComponent();
        Loaded += MainWindow_Loaded;
    }

    private async void MainWindow_Loaded(object sender, RoutedEventArgs e)
    {
        await InitializeWebViewAsync();

        if (_tabs.Count == 0)
        {
            CreateNewTab();
        }
    }

    private async Task InitializeWebViewAsync()
    {
        if (_isWebViewInitialized) return;

        try
        {
            var env = await CoreWebView2Environment.CreateAsync(userDataFolder: Path.Combine(Path.GetTempPath(), "PdfRead_WebView2"));
            await MainWebView.EnsureCoreWebView2Async(env);
            _isWebViewInitialized = true;

            // Configure Edge PDF viewer settings
            MainWebView.CoreWebView2.Settings.IsStatusBarEnabled = false;
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to initialize WebView2: {ex.Message}", "WebView2 Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public void OpenTarget(string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return;

        Dispatcher.Invoke(async () =>
        {
            try
            {
                string ext = Path.GetExtension(path).ToLowerInvariant();

                // If Office file, automatically convert to PDF first
                if (ext is ".docx" or ".doc" or ".xlsx" or ".xls" or ".pptx" or ".ppt")
                {
                    TxtStatus.Text = $"Converting {Path.GetFileName(path)} to PDF via Microsoft Office...";
                    string convertedPdf = await _officeConverter.ConvertToPdfAsync(path);
                    path = convertedPdf;
                }

                if (!File.Exists(path)) return;

                // Check if already open in a tab
                var existingTab = _tabs.FirstOrDefault(t => string.Equals(t.FilePath, path, StringComparison.OrdinalIgnoreCase));
                if (existingTab != null)
                {
                    SelectTab(existingTab);
                    return;
                }

                // If active tab is untitled and unmodified, reuse it
                PdfDocumentTab targetTab;
                if (_activeTab != null && string.IsNullOrEmpty(_activeTab.FilePath) && !_activeTab.IsModified)
                {
                    targetTab = _activeTab;
                }
                else
                {
                    targetTab = CreateNewTab();
                }

                targetTab.FilePath = path;
                SelectTab(targetTab);
                await LoadPdfInTabAsync(targetTab);
                TxtStatus.Text = $"Loaded: {Path.GetFileName(path)}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error opening document: {ex.Message}", "Open Error", MessageBoxButton.OK, MessageBoxImage.Error);
                TxtStatus.Text = "Error opening file.";
            }
        });
    }

    private PdfDocumentTab CreateNewTab(string? filePath = null)
    {
        var tab = new PdfDocumentTab
        {
            FilePath = filePath ?? string.Empty
        };

        _tabs.Add(tab);
        RebuildTabsUi();
        SelectTab(tab);

        return tab;
    }

    private void SelectTab(PdfDocumentTab tab)
    {
        _activeTab = tab;
        RebuildTabsUi();

        if (!string.IsNullOrEmpty(tab.FilePath) && File.Exists(tab.FilePath))
        {
            _ = LoadPdfInTabAsync(tab);
            TxtDocInfo.Text = tab.FilePath;
        }
        else
        {
            TxtDocInfo.Text = "Untitled Document";
        }
    }

    private void CloseTab(PdfDocumentTab tab)
    {
        int index = _tabs.IndexOf(tab);
        _tabs.Remove(tab);

        if (_tabs.Count == 0)
        {
            CreateNewTab();
        }
        else
        {
            int nextIndex = Math.Clamp(index - 1, 0, _tabs.Count - 1);
            SelectTab(_tabs[nextIndex]);
        }
    }

    private void RebuildTabsUi()
    {
        // Keep BtnNewTab at the end
        TabsContainer.Children.Clear();

        foreach (var tab in _tabs)
        {
            bool isActive = tab == _activeTab;
            var btnStyle = (Style)FindResource(isActive ? "ActiveTabItemButtonStyle" : "TabItemButtonStyle");

            var tabBtn = new Button
            {
                Style = btnStyle,
                Tag = tab
            };

            var panel = new StackPanel { Orientation = Orientation.Horizontal };
            var titleBlock = new TextBlock
            {
                Text = tab.Title + (tab.IsModified ? " •" : ""),
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(0, 0, 8, 0)
            };

            var closeBtn = new Button
            {
                Content = "×",
                Background = Brushes.Transparent,
                Foreground = Brushes.Gray,
                BorderThickness = new Thickness(0),
                Cursor = Cursors.Hand,
                Padding = new Thickness(4, 0, 4, 0),
                FontWeight = FontWeights.Bold,
                FontSize = 13
            };

            closeBtn.Click += (s, e) =>
            {
                e.Handled = true;
                CloseTab(tab);
            };

            panel.Children.Add(titleBlock);
            panel.Children.Add(closeBtn);
            tabBtn.Content = panel;

            tabBtn.Click += (s, e) => SelectTab(tab);
            TabsContainer.Children.Add(tabBtn);
        }

        TabsContainer.Children.Add(BtnNewTab);
    }

    private async Task LoadPdfInTabAsync(PdfDocumentTab tab)
    {
        if (string.IsNullOrEmpty(tab.FilePath) || !File.Exists(tab.FilePath)) return;

        await InitializeWebViewAsync();

        if (_isWebViewInitialized)
        {
            // Navigate directly to the local PDF file path using file:// protocol
            string fileUri = new Uri(tab.FilePath).AbsoluteUri;
            MainWebView.CoreWebView2.Navigate(fileUri);
        }
    }

    // ----------------- TOOLBAR BUTTON HANDLERS -----------------

    private void BtnNewTab_Click(object sender, RoutedEventArgs e)
    {
        CreateNewTab();
    }

    private void BtnOpen_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter = "All Supported Files (*.pdf;*.docx;*.xlsx;*.pptx)|*.pdf;*.docx;*.xlsx;*.pptx|PDF Files (*.pdf)|*.pdf|Word Documents (*.docx;*.doc)|*.docx;*.doc|Excel Spreadsheets (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
            Title = "Open PDF or Document"
        };

        if (dlg.ShowDialog(this) == true)
        {
            OpenTarget(dlg.FileName);
        }
    }

    private async void BtnSave_Click(object sender, RoutedEventArgs e)
    {
        if (_activeTab == null || string.IsNullOrEmpty(_activeTab.FilePath)) return;

        try
        {
            var modified = _currentTextBlocks.Where(b => b.IsModified).ToList();
            if (modified.Count > 0)
            {
                TxtStatus.Text = "Saving in-place text edits to PDF...";
                string tempOut = Path.Combine(Path.GetTempPath(), $"saved_{Guid.NewGuid():N}.pdf");
                await _editorService.SaveTextModificationsAsync(_activeTab.FilePath, tempOut, modified);

                File.Copy(tempOut, _activeTab.FilePath, true);
                File.Delete(tempOut);

                _activeTab.IsModified = false;
                RebuildTabsUi();
                await LoadPdfInTabAsync(_activeTab);
                TxtStatus.Text = "Document saved successfully!";
            }
            else
            {
                TxtStatus.Text = "Document is already up-to-date.";
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to save document: {ex.Message}", "Save Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BtnTextEditMode_Click(object sender, RoutedEventArgs e)
    {
        if (_activeTab == null || string.IsNullOrEmpty(_activeTab.FilePath))
        {
            MessageBox.Show("Please open a PDF document first.", "No Document", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        bool isVisible = TextEditOverlayPanel.Visibility == Visibility.Visible;
        TextEditOverlayPanel.Visibility = isVisible ? Visibility.Collapsed : Visibility.Visible;

        if (!isVisible)
        {
            _ = ScanCurrentPageText();
        }
    }

    private void BtnCloseTextEdit_Click(object sender, RoutedEventArgs e)
    {
        TextEditOverlayPanel.Visibility = Visibility.Collapsed;
    }

    private async void BtnScanPageText_Click(object sender, RoutedEventArgs e)
    {
        await ScanCurrentPageText();
    }

    private async Task ScanCurrentPageText()
    {
        if (_activeTab == null || string.IsNullOrEmpty(_activeTab.FilePath)) return;

        try
        {
            TxtStatus.Text = "Scanning page text blocks...";
            var blocks = await _editorService.ExtractTextBlocksAsync(_activeTab.FilePath, _activeTab.CurrentPage);
            _currentTextBlocks.Clear();
            _currentTextBlocks.AddRange(blocks);

            LstTextBlocks.ItemsSource = null;
            LstTextBlocks.ItemsSource = _currentTextBlocks;
            TxtStatus.Text = $"Extracted {_currentTextBlocks.Count} text elements on page {_activeTab.CurrentPage}.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error extracting text: {ex.Message}", "Scan Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void LstTextBlocks_SelectionChanged(object sender, SelectionChangedEventArgs e)
    {
        if (LstTextBlocks.SelectedItem is TextBlockItem item)
        {
            PnlBlockEditor.Visibility = Visibility.Visible;
            TxtBlockContent.Text = item.Text;
            TxtBlockFontSize.Text = item.FontSize.ToString("0.#");
            TxtBlockFontFamily.Text = item.FontFamily;
        }
        else
        {
            PnlBlockEditor.Visibility = Visibility.Collapsed;
        }
    }

    private void BtnApplyBlockEdit_Click(object sender, RoutedEventArgs e)
    {
        if (LstTextBlocks.SelectedItem is TextBlockItem item && _activeTab != null)
        {
            item.Text = TxtBlockContent.Text;
            if (double.TryParse(TxtBlockFontSize.Text, out double fSize)) item.FontSize = fSize;
            if (!string.IsNullOrWhiteSpace(TxtBlockFontFamily.Text)) item.FontFamily = TxtBlockFontFamily.Text.Trim();
            item.IsModified = true;

            _activeTab.IsModified = true;
            RebuildTabsUi();
            LstTextBlocks.Items.Refresh();
            TxtStatus.Text = "Text edit applied! Click 💾 Save to write updates to PDF.";
        }
    }

    private async void BtnRotate_Click(object sender, RoutedEventArgs e)
    {
        if (_activeTab == null || string.IsNullOrEmpty(_activeTab.FilePath)) return;

        try
        {
            TxtStatus.Text = "Rotating page...";
            await _editorService.RotatePageAsync(_activeTab.FilePath, _activeTab.CurrentPage, 90);
            await LoadPdfInTabAsync(_activeTab);
            TxtStatus.Text = "Page rotated 90° clockwise.";
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Rotate error: {ex.Message}", "Rotate Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void BtnDeletePage_Click(object sender, RoutedEventArgs e)
    {
        if (_activeTab == null || string.IsNullOrEmpty(_activeTab.FilePath)) return;

        if (MessageBox.Show("Delete current page from document?", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
        {
            try
            {
                TxtStatus.Text = "Deleting page...";
                await _editorService.DeletePageAsync(_activeTab.FilePath, _activeTab.CurrentPage);
                await LoadPdfInTabAsync(_activeTab);
                TxtStatus.Text = "Page deleted.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Delete error: {ex.Message}", "Delete Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private async void BtnExportExcel_Click(object sender, RoutedEventArgs e)
    {
        if (_activeTab == null || string.IsNullOrEmpty(_activeTab.FilePath))
        {
            MessageBox.Show("Please open a PDF document first.", "No Document", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dlg = new SaveFileDialog
        {
            Filter = "Excel Spreadsheet (*.xlsx)|*.xlsx",
            FileName = Path.ChangeExtension(Path.GetFileName(_activeTab.FilePath), ".xlsx")
        };

        if (dlg.ShowDialog(this) == true)
        {
            try
            {
                TxtStatus.Text = "Extracting tables and generating Excel spreadsheet...";
                string outPath = await _exportService.ExportToExcelAsync(_activeTab.FilePath, dlg.FileName);
                TxtStatus.Text = $"Exported to Excel: {Path.GetFileName(outPath)}";
                MessageBox.Show($"Spreadsheet successfully exported to:\n{outPath}", "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Excel export failed: {ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private async void BtnExportWord_Click(object sender, RoutedEventArgs e)
    {
        if (_activeTab == null || string.IsNullOrEmpty(_activeTab.FilePath))
        {
            MessageBox.Show("Please open a PDF document first.", "No Document", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        var dlg = new SaveFileDialog
        {
            Filter = "Word Document (*.docx)|*.docx",
            FileName = Path.ChangeExtension(Path.GetFileName(_activeTab.FilePath), ".docx")
        };

        if (dlg.ShowDialog(this) == true)
        {
            try
            {
                TxtStatus.Text = "Converting document to Word format...";
                string outPath = await _exportService.ExportToWordAsync(_activeTab.FilePath, dlg.FileName);
                TxtStatus.Text = $"Exported to Word: {Path.GetFileName(outPath)}";
                MessageBox.Show($"Word document successfully exported to:\n{outPath}", "Export Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Word export failed: {ex.Message}", "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private async void BtnConvertOffice_Click(object sender, RoutedEventArgs e)
    {
        var dlg = new OpenFileDialog
        {
            Filter = "Office Documents (*.docx;*.doc;*.xlsx;*.xls;*.pptx;*.ppt)|*.docx;*.doc;*.xlsx;*.xls;*.pptx;*.ppt|All Files (*.*)|*.*",
            Title = "Select Word or Excel File to Convert to PDF"
        };

        if (dlg.ShowDialog(this) == true)
        {
            try
            {
                TxtStatus.Text = $"Converting {Path.GetFileName(dlg.FileName)} via Microsoft Office...";
                string pdfPath = await _officeConverter.ConvertToPdfAsync(dlg.FileName);
                TxtStatus.Text = $"Converted: {Path.GetFileName(pdfPath)}";
                OpenTarget(pdfPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Office conversion failed: {ex.Message}", "Conversion Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }

    private void BtnSearch_Click(object sender, RoutedEventArgs e)
    {
        TriggerSearch();
    }

    private void TxtSearch_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.Key == Key.Enter)
        {
            TriggerSearch();
        }
    }

    private void TriggerSearch()
    {
        // Edge WebView2 native PDF viewer supports Ctrl+F
        MainWebView.Focus();
    }

    private void BtnIntegration_Click(object sender, RoutedEventArgs e)
    {
        var settingsWin = new SettingsWindow { Owner = this };
        settingsWin.ShowDialog();
    }

    // ----------------- DRAG & DROP & KEYBOARD -----------------

    private void Window_PreviewDragOver(object sender, DragEventArgs e)
    {
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            e.Effects = DragDropEffects.Copy;
            DropOverlay.Visibility = Visibility.Visible;
            e.Handled = true;
        }
    }

    private void Window_Drop(object sender, DragEventArgs e)
    {
        DropOverlay.Visibility = Visibility.Collapsed;
        if (e.Data.GetDataPresent(DataFormats.FileDrop))
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (var file in files)
            {
                OpenTarget(file);
            }
        }
    }

    private void Window_KeyDown(object sender, KeyEventArgs e)
    {
        if (Keyboard.Modifiers == ModifierKeys.Control)
        {
            switch (e.Key)
            {
                case Key.O:
                    BtnOpen_Click(this, new RoutedEventArgs());
                    e.Handled = true;
                    break;
                case Key.S:
                    BtnSave_Click(this, new RoutedEventArgs());
                    e.Handled = true;
                    break;
                case Key.T:
                    CreateNewTab();
                    e.Handled = true;
                    break;
                case Key.W:
                    if (_activeTab != null) CloseTab(_activeTab);
                    e.Handled = true;
                    break;
                case Key.F:
                    TxtSearch.Focus();
                    e.Handled = true;
                    break;
            }
        }
    }
}
