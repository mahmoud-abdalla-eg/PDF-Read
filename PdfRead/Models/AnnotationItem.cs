using System;

namespace PdfRead.Models;

public class AnnotationItem
{
    public string Id { get; set; } = Guid.NewGuid().ToString("N");
    public int PageIndex { get; set; } = 1;
    public string Type { get; set; } = "highlight"; // "highlight", "note", "ink", "stamp", "rect"
    public double X { get; set; }
    public double Y { get; set; }
    public double Width { get; set; }
    public double Height { get; set; }
    public string Color { get; set; } = "#FFE500";
    public double Opacity { get; set; } = 0.4;
    public string Content { get; set; } = string.Empty;
}
