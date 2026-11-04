using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using UglyToad.PdfPig;

namespace PdfRead.Services;

public class DocumentExportService
{
    private static string? _cachedPythonExe;

    public async Task<string> ExportToWordAsync(string pdfPath, string? outputDocxPath = null)
    {
        if (!File.Exists(pdfPath))
        {
            throw new FileNotFoundException("PDF document not found", pdfPath);
        }

        if (string.IsNullOrEmpty(outputDocxPath))
        {
            outputDocxPath = Path.ChangeExtension(pdfPath, ".docx");
        }

        string resolvedOut = Path.GetFullPath(outputDocxPath);

        // 1. Try High-Fidelity Python Engine (pdf2docx)
        bool pythonSuccess = await TryConvertWithPythonScriptAsync("convert_pdf_to_word.py", pdfPath, resolvedOut);
        if (pythonSuccess && File.Exists(resolvedOut) && new FileInfo(resolvedOut).Length > 100)
        {
            return resolvedOut;
        }

        // 2. Fallback: Pure C# DocumentFormat.OpenXml Generator
        await Task.Run(() =>
        {
            using var wordDoc = WordprocessingDocument.Create(resolvedOut, WordprocessingDocumentType.Document);
            var mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            var body = mainPart.Document.AppendChild(new Body());

            using var pdf = UglyToad.PdfPig.PdfDocument.Open(pdfPath);
            foreach (var page in pdf.GetPages())
            {
                var words = page.GetWords().OrderByDescending(w => w.BoundingBox.Top).ThenBy(w => w.BoundingBox.Left).ToList();
                if (words.Count == 0) continue;

                // Group words into lines
                var lines = new List<List<UglyToad.PdfPig.Content.Word>>();
                var currentLine = new List<UglyToad.PdfPig.Content.Word>();
                double lastY = double.NaN;

                foreach (var word in words)
                {
                    double wordY = word.BoundingBox.Bottom;
                    if (double.IsNaN(lastY) || Math.Abs(wordY - lastY) <= 4.0)
                    {
                        currentLine.Add(word);
                        if (double.IsNaN(lastY)) lastY = wordY;
                    }
                    else
                    {
                        if (currentLine.Count > 0)
                        {
                            lines.Add(currentLine.OrderBy(w => w.BoundingBox.Left).ToList());
                        }
                        currentLine = new List<UglyToad.PdfPig.Content.Word> { word };
                        lastY = wordY;
                    }
                }
                if (currentLine.Count > 0) lines.Add(currentLine.OrderBy(w => w.BoundingBox.Left).ToList());

                foreach (var line in lines)
                {
                    string lineText = string.Join(" ", line.Select(w => w.Text));
                    var p = body.AppendChild(new Paragraph());
                    var r = p.AppendChild(new Run());
                    r.AppendChild(new Text(lineText));
                }

                // Page break between pages
                var breakP = body.AppendChild(new Paragraph());
                var breakR = breakP.AppendChild(new Run());
                breakR.AppendChild(new Break { Type = BreakValues.Page });
            }

            mainPart.Document.Save();
        });

        return resolvedOut;
    }

    public async Task<string> ExportToExcelAsync(string pdfPath, string? outputXlsxPath = null)
    {
        if (!File.Exists(pdfPath))
        {
            throw new FileNotFoundException("PDF document not found", pdfPath);
        }

        if (string.IsNullOrEmpty(outputXlsxPath))
        {
            outputXlsxPath = Path.ChangeExtension(pdfPath, ".xlsx");
        }

        string resolvedOut = Path.GetFullPath(outputXlsxPath);

        // 1. Try High-Fidelity Python Engine (pdfplumber + openpyxl)
        bool pythonSuccess = await TryConvertWithPythonScriptAsync("convert_pdf_to_excel.py", pdfPath, resolvedOut);
        if (pythonSuccess && File.Exists(resolvedOut) && new FileInfo(resolvedOut).Length > 100)
        {
            return resolvedOut;
        }

        // 2. Fallback: Pure C# ClosedXML Engine with styled tables and RTL detection
        await Task.Run(() =>
        {
            using var workbook = new XLWorkbook();
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(pdfPath);

            int pageNum = 1;
            foreach (var page in pdf.GetPages())
            {
                string sheetName = $"Page {pageNum}";
                var worksheet = workbook.Worksheets.Add(sheetName);

                bool isArabic = page.GetWords().Any(w => System.Text.RegularExpressions.Regex.IsMatch(w.Text, @"[\u0600-\u06FF\u0750-\u077F\u08A0-\u08FF\uFB50-\uFDFF\uFE70-\uFEFF]"));
                if (isArabic)
                {
                    worksheet.RightToLeft = true;
                }

                var words = page.GetWords().OrderByDescending(w => w.BoundingBox.Top).ToList();
                if (words.Count == 0)
                {
                    pageNum++;
                    continue;
                }

                var lines = new List<List<UglyToad.PdfPig.Content.Word>>();
                var currentLine = new List<UglyToad.PdfPig.Content.Word>();
                double lastY = double.NaN;

                foreach (var word in words)
                {
                    double wordY = word.BoundingBox.Bottom;
                    if (double.IsNaN(lastY) || Math.Abs(wordY - lastY) <= 4.0)
                    {
                        currentLine.Add(word);
                        if (double.IsNaN(lastY)) lastY = wordY;
                    }
                    else
                    {
                        if (currentLine.Count > 0)
                        {
                            lines.Add(currentLine.OrderBy(w => w.BoundingBox.Left).ToList());
                        }
                        currentLine = new List<UglyToad.PdfPig.Content.Word> { word };
                        lastY = wordY;
                    }
                }
                if (currentLine.Count > 0) lines.Add(currentLine.OrderBy(w => w.BoundingBox.Left).ToList());

                int rowIndex = 1;
                foreach (var line in lines)
                {
                    var cells = new List<string>();
                    var currentCellWords = new List<string>();
                    double lastRight = double.NaN;

                    foreach (var word in line)
                    {
                        if (!double.IsNaN(lastRight) && (word.BoundingBox.Left - lastRight) > 15.0)
                        {
                            cells.Add(string.Join(" ", currentCellWords));
                            currentCellWords.Clear();
                        }

                        currentCellWords.Add(word.Text);
                        lastRight = word.BoundingBox.Right;
                    }

                    if (currentCellWords.Count > 0)
                    {
                        cells.Add(string.Join(" ", currentCellWords));
                    }

                    for (int colIndex = 0; colIndex < cells.Count; colIndex++)
                    {
                        string val = cells[colIndex].Trim();
                        var cell = worksheet.Cell(rowIndex, colIndex + 1);

                        if (double.TryParse(val, out double numVal))
                        {
                            cell.Value = numVal;
                        }
                        else
                        {
                            cell.Value = val;
                        }

                        cell.Style.Border.SetOutsideBorder(XLBorderStyleValues.Thin);
                        cell.Style.Border.SetOutsideBorderColor(XLColor.FromHtml("#CBD5E1"));

                        if (rowIndex == 1)
                        {
                            cell.Style.Fill.SetBackgroundColor(XLColor.FromHtml("#1E293B"));
                            cell.Style.Font.SetFontColor(XLColor.White);
                            cell.Style.Font.SetBold(true);
                            cell.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        }
                    }

                    rowIndex++;
                }

                worksheet.Columns().AdjustToContents(1, 40);
                pageNum++;
            }

            workbook.SaveAs(resolvedOut);
        });

        return resolvedOut;
    }

    private static async Task<bool> TryConvertWithPythonScriptAsync(string scriptName, string inputPdf, string outputTarget)
    {
        string? pythonExe = GetPythonExecutable();
        if (string.IsNullOrEmpty(pythonExe)) return false;

        string scriptPath = FindScriptPath(scriptName);
        if (!File.Exists(scriptPath)) return false;

        try
        {
            var psi = new ProcessStartInfo
            {
                FileName = pythonExe,
                Arguments = $"\"{scriptPath}\" \"{inputPdf}\" \"{outputTarget}\"",
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using var process = Process.Start(psi);
            if (process == null) return false;

            var stdoutTask = process.StandardOutput.ReadToEndAsync();
            var stderrTask = process.StandardError.ReadToEndAsync();

            var timeoutTask = Task.Delay(TimeSpan.FromSeconds(45));
            var completed = await Task.WhenAny(process.WaitForExitAsync(), timeoutTask);

            if (completed == timeoutTask)
            {
                try { process.Kill(); } catch { }
                return false;
            }

            string stdout = await stdoutTask;
            string stderr = await stderrTask;

            return process.ExitCode == 0 && File.Exists(outputTarget);
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"Python conversion exception: {ex.Message}");
            return false;
        }
    }

    private static string FindScriptPath(string scriptName)
    {
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        string inScripts = Path.Combine(baseDir, "Scripts", scriptName);
        if (File.Exists(inScripts)) return inScripts;

        string inAssets = Path.Combine(baseDir, "Assets", scriptName);
        if (File.Exists(inAssets)) return inAssets;

        string inBase = Path.Combine(baseDir, scriptName);
        if (File.Exists(inBase)) return inBase;

        // Try parent directories up to workspace
        string? parent = baseDir;
        for (int i = 0; i < 4 && parent != null; i++)
        {
            string candidate = Path.Combine(parent, "scripts", scriptName);
            if (File.Exists(candidate)) return candidate;
            string candidate2 = Path.Combine(parent, "PdfRead", "Scripts", scriptName);
            if (File.Exists(candidate2)) return candidate2;
            parent = Directory.GetParent(parent)?.FullName;
        }

        return inScripts;
    }

    private static string? GetPythonExecutable()
    {
        if (_cachedPythonExe != null) return _cachedPythonExe;

        string[] candidates = new[]
        {
            "py",
            "python",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Python", "pythoncore-3.14-64", "python.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Python", "Python314", "python.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Python", "Python313", "python.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Python", "Python312", "python.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Programs", "Python", "Python311", "python.exe")
        };

        foreach (var candidate in candidates)
        {
            try
            {
                var psi = new ProcessStartInfo
                {
                    FileName = candidate,
                    Arguments = "--version",
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using var p = Process.Start(psi);
                if (p != null)
                {
                    p.WaitForExit(1500);
                    if (p.ExitCode == 0)
                    {
                        _cachedPythonExe = candidate;
                        return candidate;
                    }
                }
            }
            catch { }
        }

        return null;
    }
}
