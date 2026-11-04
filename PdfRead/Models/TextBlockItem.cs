namespace PdfRead.Models;

public class TextBlockItem
{
    public int PageIndex { get; set; } = 1;
    public double X { get; set; }
    public double Y { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
    public string Text { get; set; } = string.Empty;
    public string OriginalText { get; set; } = string.Empty;
    public string FontFamily { get; set; } = "Arial";
    public double FontSize { get; set; } = 12.0;
    public string Color { get; set; } = "#000000";
    public bool IsModified { get; set; }
}
