using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using PdfRead.Models;
using PdfSharp.Drawing;
using PdfSharp.Fonts;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using UglyToad.PdfPig;

namespace PdfRead.Services;

public class PdfEditorService
{
    static PdfEditorService()
    {
        try
        {
            GlobalFontSettings.FontResolver ??= new WindowsFontResolver();
        }
        catch { }
    }
    public Task<List<TextBlockItem>> ExtractTextBlocksAsync(string pdfPath, int pageIndex)
    {
        return Task.Run(() =>
        {
            var results = new List<TextBlockItem>();
            if (!File.Exists(pdfPath)) return results;

            using var pdf = UglyToad.PdfPig.PdfDocument.Open(pdfPath);
            if (pageIndex < 1 || pageIndex > pdf.NumberOfPages) return results;

            var page = pdf.GetPage(pageIndex);
            double pageHeight = page.Height;

            // Group letters into words and words into lines
            var words = page.GetWords().OrderByDescending(w => w.BoundingBox.Top).ToList();
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

            if (currentLine.Count > 0)
            {
                lines.Add(currentLine.OrderBy(w => w.BoundingBox.Left).ToList());
            }

            foreach (var line in lines)
            {
                if (line.Count == 0) continue;

                double minX = line.Min(w => w.BoundingBox.Left);
                double maxX = line.Max(w => w.BoundingBox.Right);
                double minY = line.Min(w => w.BoundingBox.Bottom);
                double maxY = line.Max(w => w.BoundingBox.Top);

                // Convert PDF coordinate (bottom-left origin) to screen coordinate (top-left origin)
                double screenY = pageHeight - maxY;
                double width = maxX - minX;
                double height = maxY - minY;

                string lineText = string.Join(" ", line.Select(w => w.Text));
                var firstWord = line[0];
                var sampleLetter = firstWord.Letters.FirstOrDefault();

                string fontName = sampleLetter?.FontName ?? "Arial";
                // Normalize font names
                if (fontName.Contains("Calibri", StringComparison.OrdinalIgnoreCase)) fontName = "Calibri";
                else if (fontName.Contains("Times", StringComparison.OrdinalIgnoreCase)) fontName = "Times New Roman";
                else if (fontName.Contains("Courier", StringComparison.OrdinalIgnoreCase)) fontName = "Courier New";
                else fontName = "Arial";

                double fontSize = sampleLetter?.FontSize ?? 12.0;

                results.Add(new TextBlockItem
                {
                    PageIndex = pageIndex,
                    X = minX,
                    Y = screenY,
                    Width = width,
                    Height = height,
                    Text = lineText,
                    OriginalText = lineText,
                    FontFamily = fontName,
                    FontSize = fontSize,
                    Color = "#000000"
                });
            }

            return results;
        });
    }

    public Task SaveTextModificationsAsync(string pdfPath, string outputPath, List<TextBlockItem> modifiedBlocks)
    {
        return Task.Run(() =>
        {
            if (!File.Exists(pdfPath))
            {
                throw new FileNotFoundException("PDF file not found", pdfPath);
            }

            // Open existing PDF with PdfSharp
            using var document = PdfSharp.Pdf.IO.PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify);

            var groupedByPage = modifiedBlocks.Where(b => b.IsModified).GroupBy(b => b.PageIndex);

            foreach (var pageGroup in groupedByPage)
            {
                int pageIdx = pageGroup.Key - 1;
                if (pageIdx < 0 || pageIdx >= document.PageCount) continue;

                var page = document.Pages[pageIdx];
                double pageHeight = page.Height.Point;

                using var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);

                foreach (var block in pageGroup)
                {
                    // Calculate PDF coordinate (bottom-left to top-left)
                    // Bounding box in top-left screen coordinates: block.X, block.Y
                    var rect = new XRect(block.X, block.Y, Math.Max(block.Width, 10), Math.Max(block.Height, 10));

                    // 1. Redact original text area with opaque white background
                    gfx.DrawRectangle(XBrushes.White, rect);

                    // 2. Draw replacement text
                    var font = new XFont(block.FontFamily, Math.Max(block.FontSize, 6.0), XFontStyleEx.Regular);
                    var textFormat = new XStringFormat
                    {
                        Alignment = XStringAlignment.Near,
                        LineAlignment = XLineAlignment.Near
                    };

                    XColor textColor = XColors.Black;
                    try
                    {
                        if (!string.IsNullOrEmpty(block.Color) && block.Color.StartsWith("#"))
                        {
                            textColor = XColor.FromArgb(
                                255,
                                Convert.ToByte(block.Color.Substring(1, 2), 16),
                                Convert.ToByte(block.Color.Substring(3, 2), 16),
                                Convert.ToByte(block.Color.Substring(5, 2), 16)
                            );
                        }
                    }
                    catch { }

                    var brush = new XSolidBrush(textColor);
                    gfx.DrawString(block.Text, font, brush, rect, textFormat);
                }
            }

            document.Save(outputPath);
        });
    }

    public Task RotatePageAsync(string pdfPath, int pageIndex, int degrees = 90)
    {
        return Task.Run(() =>
        {
            using var document = PdfSharp.Pdf.IO.PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify);
            int idx = pageIndex - 1;
            if (idx >= 0 && idx < document.PageCount)
            {
                var page = document.Pages[idx];
                page.Rotate = (page.Rotate + degrees) % 360;
                document.Save(pdfPath);
            }
        });
    }

    public Task DeletePageAsync(string pdfPath, int pageIndex)
    {
        return Task.Run(() =>
        {
            using var document = PdfSharp.Pdf.IO.PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify);
            int idx = pageIndex - 1;
            if (idx >= 0 && idx < document.PageCount && document.PageCount > 1)
            {
                document.Pages.RemoveAt(idx);
                document.Save(pdfPath);
            }
        });
    }
}
