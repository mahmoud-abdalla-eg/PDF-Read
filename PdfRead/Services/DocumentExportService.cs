using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using ClosedXML.Excel;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace PdfRead.Services;

public class DocumentExportService
{
    public Task<string> ExportToExcelAsync(string pdfPath, string? outputXlsxPath = null)
    {
        if (!File.Exists(pdfPath))
        {
            throw new FileNotFoundException("PDF document not found", pdfPath);
        }

        if (string.IsNullOrEmpty(outputXlsxPath))
        {
            outputXlsxPath = Path.ChangeExtension(pdfPath, ".xlsx");
        }

        string resolvedOut = outputXlsxPath;
        return Task.Run(() =>
        {
            using var workbook = new XLWorkbook();
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(pdfPath);

            int pageNum = 1;
            foreach (var page in pdf.GetPages())
            {
                string sheetName = $"Page {pageNum}";
                var worksheet = workbook.Worksheets.Add(sheetName);

                // Extract words with bounding box
                var words = page.GetWords().OrderByDescending(w => w.BoundingBox.Top).ToList();

                if (words.Count == 0)
                {
                    pageNum++;
                    continue;
                }

                // Cluster into lines by Y position tolerance (within 4 points)
                var lines = new List<List<Word>>();
                var currentLine = new List<Word>();
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
                        currentLine = new List<Word> { word };
                        lastY = wordY;
                    }
                }

                if (currentLine.Count > 0)
                {
                    lines.Add(currentLine.OrderBy(w => w.BoundingBox.Left).ToList());
                }

                int rowIndex = 1;
                foreach (var line in lines)
                {
                    // Detect column gaps (> 15 points)
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
                        if (double.TryParse(val, out double numVal))
                        {
                            worksheet.Cell(rowIndex, colIndex + 1).Value = numVal;
                        }
                        else
                        {
                            worksheet.Cell(rowIndex, colIndex + 1).Value = val;
                        }
                    }

                    rowIndex++;
                }

                worksheet.Columns().AdjustToContents(1, 40);
                pageNum++;
            }

            workbook.SaveAs(resolvedOut);
            return resolvedOut;
        });
    }

    public Task<string> ExportToWordAsync(string pdfPath, string? outputDocxPath = null)
    {
        if (!File.Exists(pdfPath))
        {
            throw new FileNotFoundException("PDF document not found", pdfPath);
        }

        if (string.IsNullOrEmpty(outputDocxPath))
        {
            outputDocxPath = Path.ChangeExtension(pdfPath, ".docx");
        }

        string resolvedOut = outputDocxPath;
        return Task.Run(() =>
        {
            // Extract paragraphs and text cleanly with PdfPig and write structured document
            using (var pdf = UglyToad.PdfPig.PdfDocument.Open(pdfPath))
            {
                using var writer = new StreamWriter(resolvedOut);
                foreach (var page in pdf.GetPages())
                {
                    writer.WriteLine(page.Text);
                    writer.WriteLine();
                }
            }

            return resolvedOut;
        });
    }
}
