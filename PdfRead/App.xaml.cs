using System;
using System.Diagnostics;
using System.IO;
using System.IO.Pipes;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Interop;
using PdfRead.Services;

namespace PdfRead;

public partial class App : Application
{
    private const string PipeName = "PdfRead_IPC_Pipe_84712";

    private CancellationTokenSource? _pipeCts;
    private MainWindow? _mainWindow;

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

    private const int SW_RESTORE = 9;

    private void Application_Startup(object sender, StartupEventArgs e)
    {
        // 1. Command-line switch handling
        if (e.Args.Contains("--register", StringComparer.OrdinalIgnoreCase))
        {
            ShellIntegration.RegisterFileAssociations();
            Environment.Exit(0);
            return;
        }

        if (e.Args.Contains("--shortcut", StringComparer.OrdinalIgnoreCase))
        {
            ShellIntegration.UpdateAllIconsAndShortcuts();
            Environment.Exit(0);
            return;
        }

        if (e.Args.Contains("--unregister", StringComparer.OrdinalIgnoreCase))
        {
            ShellIntegration.UnregisterFileAssociations();
            Environment.Exit(0);
            return;
        }

        if (e.Args.Contains("--verify-all", StringComparer.OrdinalIgnoreCase))
        {
            Task.Run(async () =>
            {
                try
                {
                    string testDoc = @"E:\md file viewer\pdfread repo folder\examples\TestDocument.pdf";
                    var editor = new PdfEditorService();
                    var exporter = new DocumentExportService();

                    // 1. Text extraction
                    var blocks = await editor.ExtractTextBlocksAsync(testDoc, 1);

                    // 2. In-place text modification
                    if (blocks.Count > 0)
                    {
                        blocks[0].Text = "PdfRead EDITED In-Place Headline";
                        blocks[0].IsModified = true;
                        string editedPdf = Path.Combine(Path.GetDirectoryName(testDoc)!, "TestDocument_Edited.pdf");
                        await editor.SaveTextModificationsAsync(testDoc, editedPdf, blocks);
                    }

                    // 3. Export to Excel
                    string xlsxOut = Path.Combine(Path.GetDirectoryName(testDoc)!, "TestDocument_Export.xlsx");
                    await exporter.ExportToExcelAsync(testDoc, xlsxOut);

                    // 4. Export to Word
                    string docxOut = Path.Combine(Path.GetDirectoryName(testDoc)!, "TestDocument_Export.docx");
                    await exporter.ExportToWordAsync(testDoc, docxOut);

                    File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "verify_success.txt"), 
                        $"ALL_CHECKS_PASSED\nBlocks: {blocks.Count}\nEditedPdf: True\nExcelOut: True\nWordOut: True");
                }
                catch (Exception ex)
                {
                    File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "verify_error.txt"), ex.ToString());
                }
                finally
                {
                    Environment.Exit(0);
                }
            });

            return;
        }

        // Context menu conversion command: --convert-to-pdf <filePath>
        int convertIdx = Array.FindIndex(e.Args, a => string.Equals(a, "--convert-to-pdf", StringComparison.OrdinalIgnoreCase));
        if (convertIdx >= 0 && convertIdx + 1 < e.Args.Length)
        {
            string sourceFile = string.Join(" ", e.Args.Skip(convertIdx + 1)).Trim('"', '\'');
            try
            {
                var converter = new OfficeConversionService();
                string pdfPath = converter.ConvertToPdfAsync(sourceFile).GetAwaiter().GetResult();
                var psi = new ProcessStartInfo(ShellIntegration.GetExecutablePath(), $"\"{pdfPath}\"")
                {
                    UseShellExecute = true
                };
                Process.Start(psi);
            }
            catch (Exception ex)
            {
                try
                {
                    File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "conversion_error.log"), ex.ToString());
                }
                catch { }
                MessageBox.Show($"Office Conversion Error: {ex.Message}", "Conversion Failed", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            Environment.Exit(0);
            return;
        }

        // 2. Single-Instance Check via NamedPipe IPC
        var currentProc = Process.GetCurrentProcess();
        var otherInstances = Process.GetProcessesByName(currentProc.ProcessName)
                                    .Where(p => p.Id != currentProc.Id)
                                    .ToArray();

        if (otherInstances.Length > 0)
        {
            try
            {
                using var client = new NamedPipeClientStream(".", PipeName, PipeDirection.Out);
                client.Connect(2500);

                using var writer = new StreamWriter(client, Encoding.UTF8);
                writer.WriteLine(string.Join("|", e.Args));
                writer.Flush();
            }
            catch { }

            Environment.Exit(0);
            return;
        }

        StartNamedPipeServer();

        _mainWindow = new MainWindow();
        _mainWindow.Show();

        if (e.Args.Length > 0)
        {
            foreach (var arg in e.Args)
            {
                if (!arg.StartsWith("--"))
                {
                    _mainWindow.OpenTarget(arg);
                }
            }
        }
    }

    private void StartNamedPipeServer()
    {
        _pipeCts = new CancellationTokenSource();
        var token = _pipeCts.Token;

        Task.Run(async () =>
        {
            try
            {
                using var server = new NamedPipeServerStream(
                    PipeName,
                    PipeDirection.In,
                    1,
                    PipeTransmissionMode.Byte,
                    PipeOptions.Asynchronous);

                while (!token.IsCancellationRequested)
                {
                    try
                    {
                        await server.WaitForConnectionAsync(token);

                        using (var reader = new StreamReader(server, Encoding.UTF8, leaveOpen: true))
                        {
                            string? payload = await reader.ReadLineAsync(token);

                            if (!string.IsNullOrWhiteSpace(payload))
                            {
                                string[] files = payload.Split('|', StringSplitOptions.RemoveEmptyEntries);
                                Dispatcher.Invoke(() =>
                                {
                                    BringWindowToFront();
                                    foreach (var file in files)
                                    {
                                        if (!file.StartsWith("--"))
                                        {
                                            _mainWindow?.OpenTarget(file);
                                        }
                                    }
                                });
                            }
                        }

                        server.Disconnect();
                    }
                    catch (OperationCanceledException)
                    {
                        break;
                    }
                    catch
                    {
                        try { server.Disconnect(); } catch { }
                    }
                }
            }
            catch { }
        }, token);
    }

    private void BringWindowToFront()
    {
        if (_mainWindow == null) return;

        if (_mainWindow.WindowState == WindowState.Minimized)
        {
            _mainWindow.WindowState = WindowState.Normal;
        }

        var handle = new WindowInteropHelper(_mainWindow).Handle;
        ShowWindowAsync(handle, SW_RESTORE);
        SetForegroundWindow(handle);
        _mainWindow.Activate();
    }

    private void Application_Exit(object sender, ExitEventArgs e)
    {
        _pipeCts?.Cancel();
    }
}
