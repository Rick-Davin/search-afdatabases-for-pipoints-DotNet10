namespace Search_AFDatabases_for_PIPoints.Models;

internal record AppFeatures
{
    public required string AssetServer { get; set; }
    public required string DatabaseSearchPattern { get; set; } = "*";
    public required string OutputFolderPath { get; set; }   
    public int AFSearchPageSize { get; set; } = 1000;
    public bool ShowRelativePIPoints { get; set; } = false;
    public bool ShowCurrentValue { get; set; } = false;
    public bool ShowFirstRecorded { get; set; } = false;   // WARNING: this is a sluggish AFSDK call
    public string? OutputWriter { get; set; } = null;    
    public bool UseReportAutoSave { get; set; } = true;
    public int AutoSaveSeconds { get; set; } = -1;
    public List<string>? TagGroupingSeparators { get; set; } = null;

}
