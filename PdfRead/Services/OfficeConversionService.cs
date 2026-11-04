using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace PdfRead.Services;

public class OfficeConversionService
{
    // WdExportFormat.wdExportFormatPDF = 17
    private const int WdExportFormatPDF = 17;
    // XlFixedFormatType.xlTypePDF = 0
    private const int XlTypePDF = 0;

    public Task<string> ConvertToPdfAsync(string inputFilePath, string? outputPdfPath = null)
    {
        inputFilePath = Path.GetFullPath(inputFilePath);
        if (!File.Exists(inputFilePath))
        {
            throw new FileNotFoundException("Input document not found", inputFilePath);
        }

        string ext = Path.GetExtension(inputFilePath).ToLowerInvariant();

        if (string.IsNullOrEmpty(outputPdfPath))
        {
            outputPdfPath = Path.ChangeExtension(inputFilePath, ".pdf");
        }

        string resolvedOut = outputPdfPath;
        var tcs = new TaskCompletionSource<string>();
        var thread = new System.Threading.Thread(() =>
        {
            try
            {
                switch (ext)
                {
                    case ".docx":
                    case ".doc":
                    case ".rtf":
                        ConvertWordToPdf(inputFilePath, resolvedOut);
                        break;

                    case ".xlsx":
                    case ".xls":
                    case ".csv":
                        ConvertExcelToPdf(inputFilePath, resolvedOut);
                        break;

                    case ".pptx":
                    case ".ppt":
                        ConvertPowerPointToPdf(inputFilePath, resolvedOut);
                        break;

                    default:
                        throw new NotSupportedException($"Conversion from '{ext}' to PDF is not supported.");
                }

                if (!File.Exists(resolvedOut))
                {
                    throw new InvalidOperationException("Conversion finished but output PDF file was not created.");
                }

                tcs.SetResult(resolvedOut);
            }
            catch (Exception ex)
            {
                tcs.SetException(ex);
            }
        });

        thread.SetApartmentState(System.Threading.ApartmentState.STA);
        thread.IsBackground = true;
        thread.Start();

        return tcs.Task;
    }

    private void ConvertWordToPdf(string inputPath, string outputPath)
    {
        Type? wordType = Type.GetTypeFromProgID("Word.Application");
        if (wordType == null)
        {
            throw new InvalidOperationException("Microsoft Word is not installed on this system.");
        }

        dynamic? wordApp = null;
        dynamic? doc = null;

        try
        {
            wordApp = Activator.CreateInstance(wordType);
            wordApp.Visible = false;
            wordApp.ScreenUpdating = false;

            doc = wordApp.Documents.Open(inputPath);
            doc.ExportAsFixedFormat(outputPath, WdExportFormatPDF);
        }
        finally
        {
            if (doc != null)
            {
                try { doc.Close(0); } catch { }
                Marshal.ReleaseComObject(doc);
            }

            if (wordApp != null)
            {
                try { wordApp.Quit(0); } catch { }
                Marshal.ReleaseComObject(wordApp);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private void ConvertExcelToPdf(string inputPath, string outputPath)
    {
        Type? excelType = Type.GetTypeFromProgID("Excel.Application");
        if (excelType == null)
        {
            throw new InvalidOperationException("Microsoft Excel is not installed on this system.");
        }

        dynamic? excelApp = null;
        dynamic? workbook = null;

        try
        {
            excelApp = Activator.CreateInstance(excelType);
            excelApp.Visible = false;
            excelApp.ScreenUpdating = false;
            excelApp.DisplayAlerts = false;

            workbook = excelApp.Workbooks.Open(inputPath);
            workbook.ExportAsFixedFormat(XlTypePDF, outputPath);
        }
        finally
        {
            if (workbook != null)
            {
                try { workbook.Close(false); } catch { }
                Marshal.ReleaseComObject(workbook);
            }

            if (excelApp != null)
            {
                try { excelApp.Quit(); } catch { }
                Marshal.ReleaseComObject(excelApp);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    private void ConvertPowerPointToPdf(string inputPath, string outputPath)
    {
        Type? pptType = Type.GetTypeFromProgID("PowerPoint.Application");
        if (pptType == null)
        {
            throw new InvalidOperationException("Microsoft PowerPoint is not installed on this system.");
        }

        dynamic? pptApp = null;
        dynamic? presentation = null;

        try
        {
            pptApp = Activator.CreateInstance(pptType);
            presentation = pptApp.Presentations.Open(inputPath, ReadOnly: true, WithWindow: false);
            // PpSaveAsFileType.ppSaveAsPDF = 32
            presentation.SaveAs(outputPath, 32);
        }
        finally
        {
            if (presentation != null)
            {
                try { presentation.Close(); } catch { }
                Marshal.ReleaseComObject(presentation);
            }

            if (pptApp != null)
            {
                try { pptApp.Quit(); } catch { }
                Marshal.ReleaseComObject(pptApp);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
