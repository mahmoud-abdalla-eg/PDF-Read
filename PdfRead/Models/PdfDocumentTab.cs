using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.CompilerServices;
using Microsoft.Web.WebView2.Wpf;

namespace PdfRead.Models;

public class PdfDocumentTab : INotifyPropertyChanged
{
    private string _filePath = string.Empty;
    private string _title = "Untitled.pdf";
    private bool _isModified;
    private int _currentPage = 1;
    private int _totalPages = 1;
    private double _zoom = 1.0;
    private bool _isTextEditMode;
    private bool _isAnnotateMode;

    public string Id { get; } = Guid.NewGuid().ToString("N");

    public string FilePath
    {
        get => _filePath;
        set
        {
            if (_filePath != value)
            {
                _filePath = value;
                OnPropertyChanged();
                Title = string.IsNullOrEmpty(value) ? "Untitled.pdf" : Path.GetFileName(value);
            }
        }
    }

    public string Title
    {
        get => _title;
        set
        {
            if (_title != value)
            {
                _title = value;
                OnPropertyChanged();
            }
        }
    }

    public bool IsModified
    {
        get => _isModified;
        set
        {
            if (_isModified != value)
            {
                _isModified = value;
                OnPropertyChanged();
            }
        }
    }

    public int CurrentPage
    {
        get => _currentPage;
        set
        {
            if (_currentPage != value)
            {
                _currentPage = value;
                OnPropertyChanged();
            }
        }
    }

    public int TotalPages
    {
        get => _totalPages;
        set
        {
            if (_totalPages != value)
            {
                _totalPages = value;
                OnPropertyChanged();
            }
        }
    }

    public double Zoom
    {
        get => _zoom;
        set
        {
            if (Math.Abs(_zoom - value) > 0.001)
            {
                _zoom = value;
                OnPropertyChanged();
            }
        }
    }

    public bool IsTextEditMode
    {
        get => _isTextEditMode;
        set
        {
            if (_isTextEditMode != value)
            {
                _isTextEditMode = value;
                OnPropertyChanged();
            }
        }
    }

    public bool IsAnnotateMode
    {
        get => _isAnnotateMode;
        set
        {
            if (_isAnnotateMode != value)
            {
                _isAnnotateMode = value;
                OnPropertyChanged();
            }
        }
    }

    public WebView2? WebViewControl { get; set; }

    public event PropertyChangedEventHandler? PropertyChanged;

    protected void OnPropertyChanged([CallerMemberName] string? name = null)
    {
        PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }
}
