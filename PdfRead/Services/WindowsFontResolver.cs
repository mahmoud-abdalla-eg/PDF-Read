using System;
using System.IO;
using PdfSharp.Fonts;

namespace PdfRead.Services;

public class WindowsFontResolver : IFontResolver
{
    public string DefaultFontName => "Arial";

    public byte[]? GetFont(string faceName)
    {
        string fontsFolder = Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        string fileName = faceName.ToLowerInvariant() switch
        {
            "calibri" => "calibri.ttf",
            "calibrib" => "calibrib.ttf",
            "times new roman" => "times.ttf",
            "segoe ui" => "segoeui.ttf",
            "arialb" => "arialbd.ttf",
            _ => "arial.ttf"
        };

        string fullPath = Path.Combine(fontsFolder, fileName);
        if (File.Exists(fullPath))
        {
            return File.ReadAllBytes(fullPath);
        }

        string fallback = Path.Combine(fontsFolder, "arial.ttf");
        return File.Exists(fallback) ? File.ReadAllBytes(fallback) : null;
    }

    public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        string name = familyName.ToLowerInvariant();
        if (name.Contains("calibri")) return new FontResolverInfo(isBold ? "calibrib" : "calibri");
        if (name.Contains("times")) return new FontResolverInfo("times new roman");
        if (name.Contains("segoe")) return new FontResolverInfo("segoe ui");
        return new FontResolverInfo(isBold ? "arialb" : "arial");
    }
}
