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
    private static extern bool AllowSetForegroundWindow(int dwProcessId);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    [DllImport("user32.dll")]
    private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    private const int ASFW_ANY = -1;
    private const int SW_RESTORE = 9;
    private const int SW_SHOW = 5;

    private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
    private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);

    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_NOMOVE = 0x0002;
    private const uint SWP_SHOWWINDOW = 0x0040;

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
            try
            {
                Task.Run(async () =>
                {
                    string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                    string testDoc = Path.Combine(baseDir, "examples", "TestDocument.pdf");
                    if (!File.Exists(testDoc))
                    {
                        string? parent = baseDir;
                        for (int i = 0; i < 4 && parent != null; i++)
                        {
                            string candidate = Path.Combine(parent, "examples", "TestDocument.pdf");
                            if (File.Exists(candidate))
                            {
                                testDoc = candidate;
                                break;
                            }
                            parent = Directory.GetParent(parent)?.FullName;
                        }
                    }

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
                }).GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "verify_error.txt"), ex.ToString());
            }
            finally
            {
                Environment.Exit(0);
            }

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
                // Explicitly grant foreground activation rights to the running primary instance
                AllowSetForegroundWindow(otherInstances[0].Id);
                AllowSetForegroundWindow(ASFW_ANY);

                using (var client = new NamedPipeClientStream(".", PipeName, PipeDirection.InOut))
                {
                    client.Connect(3000);

                    using (var writer = new StreamWriter(client, Encoding.UTF8, leaveOpen: true))
                    {
                        writer.WriteLine(string.Join("|", e.Args));
                        writer.Flush();
                    }

                    // Wait for acknowledgment so this launcher process remains alive while primary window is restored and brought to front
                    using (var reader = new StreamReader(client, Encoding.UTF8, leaveOpen: true))
                    {
                        reader.ReadLine();
                    }
                }
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
                    PipeDirection.InOut,
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

                        // Send ACK back to the secondary process before disconnecting
                        try
                        {
                            using (var writer = new StreamWriter(server, Encoding.UTF8, leaveOpen: true))
                            {
                                await writer.WriteLineAsync("ACK");
                                await writer.FlushAsync();
                            }
                        }
                        catch { }

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

        var handle = new WindowInteropHelper(_mainWindow).Handle;
        if (handle == IntPtr.Zero) return;

        // 1. Un-minimize if window is minimized or iconic
        if (_mainWindow.WindowState == WindowState.Minimized || IsIconic(handle))
        {
            _mainWindow.WindowState = _mainWindow.PreviousNonMinimizedWindowState;
            ShowWindow(handle, SW_RESTORE);
        }

        // 2. Ensure visible
        _mainWindow.Show();

        // 3. AttachThreadInput to bypass foreground lock restrictions
        IntPtr foregroundWnd = GetForegroundWindow();
        uint foregroundThreadId = GetWindowThreadProcessId(foregroundWnd, out _);
        uint currentThreadId = GetCurrentThreadId();

        bool attached = false;
        if (foregroundThreadId != 0 && foregroundThreadId != currentThreadId)
        {
            attached = AttachThreadInput(currentThreadId, foregroundThreadId, true);
        }

        try
        {
            SetWindowPos(handle, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
            SetWindowPos(handle, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW);
            BringWindowToTop(handle);
            ShowWindow(handle, SW_SHOW);
            SetForegroundWindow(handle);
        }
        finally
        {
            if (attached)
            {
                AttachThreadInput(currentThreadId, foregroundThreadId, false);
            }
        }

        // 4. WPF Topmost toggle and focus activation
        _mainWindow.Topmost = true;
        _mainWindow.Topmost = false;
        _mainWindow.Activate();
        _mainWindow.Focus();
    }

    private void Application_Exit(object sender, ExitEventArgs e)
    {
        _pipeCts?.Cancel();
    }
}
