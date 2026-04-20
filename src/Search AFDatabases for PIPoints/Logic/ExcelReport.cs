using ClosedXML.Excel;

namespace Search_AFDatabases_for_PIPoints.Logic;

internal class ExcelReport : ReportWriter
{
    private ExcelReport(AppFeatures config, DateTime creationTime, IXLWorkbook workbook, IXLWorksheet dataSheet) : base(config, creationTime)
    {
        this.Book = workbook;
        this.DataSheet = dataSheet;
    }

    public static ExcelReport CreateAndInitialize(AppFeatures config, DateTime? creationTime = default)
    {
        if (!creationTime.HasValue)
        {
            creationTime = DateTime.UtcNow;
        }
        var book = new XLWorkbook();
        // Quick, simple worksheet creation.
        var sheet = book.Worksheets.Add(creationTime.Value.ToString("yyyy-MM-dd HHmm"));
        var instance = new ExcelReport(config, creationTime.Value, book, sheet);
        // Not-so-quick or simple worksheet initialization.
        instance.Initialize();
        return instance;
    }


    private IXLWorkbook Book { get; }
    private IXLWorksheet DataSheet { get; }

    internal override string FileName => $"SearchResults.{CreationTime:yyyy-MM-dd}_{CreationTime:HHmm}.xlsx";
    internal override int MaximumRowCount => 1_000_000;
    internal override int IndexStartsAt => 1;

    private bool IsSheetInitialized { get; set; }

    public override void Initialize()
    {
        if (IsSheetInitialized)
        {
            return;
        }

        HeaderMap = CreateHeaderMap();
        foreach (var entry in HeaderMap)
        {
            DataSheet.Cell(row: 1, column: entry.Value).Value = entry.Key;
        }

        DataSheet.SheetView.Freeze(1, 0);
        DataSheet.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;

        var headerRange = DataSheet.Range(1, 1, 1, HeaderMap.Count);
        DrawBorders(headerRange);
        headerRange.Style.Alignment.SetWrapText(true);
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Font.FontColor = XLColor.White;
        headerRange.Style.Fill.BackgroundColor = XLColor.Blue;
        // Now we color correct the PI-related columns.
        foreach (var heading in OrderedPIHeadings!)
        {
            var col = HeaderMap[heading];
            DataSheet.Cell(1, col).Style.Fill.BackgroundColor = XLColor.Maroon;
        }

        // Set specific column widths and/or alignment.
        // The order below is NOT neccesarily the order on the data sheet.
        // AF-related Columns
        FormatColumn(Heading.AssetServer, width: 15);
        FormatColumn(Heading.Database, width: 40);
        FormatColumn(Heading.ElementPath, width: 40);
        FormatColumn(Heading.AttributePath, width: 60);
        FormatColumn(Heading.AttributeGuid, width: 38);
        FormatColumn(Heading.AttributeDescription, width: 40);
        FormatColumn(Heading.AttributeType, width: 12, horizontalAlignment: XLAlignmentHorizontalValues.Center);
        FormatColumn(Heading.DefaultUom, width: 12, horizontalAlignment: XLAlignmentHorizontalValues.Center);
        FormatColumn(Heading.SourceUom, width: 12, horizontalAlignment: XLAlignmentHorizontalValues.Center);
        FormatColumn(Heading.DataReference, width: 15, horizontalAlignment: XLAlignmentHorizontalValues.Center);
        FormatColumn(Heading.ConfigString, width: 30);
        // PI-related columns
        FormatColumn(Heading.DataArchive, width: 15);
        FormatColumn(Heading.TagGrouping, width: 15);
        FormatColumn(Heading.Tag, width: 40);
        FormatColumn(Heading.TagDescriptor, width: 40);
        FormatColumn(Heading.PointId, width: 12, horizontalAlignment: XLAlignmentHorizontalValues.Center);
        FormatColumn(Heading.PointType, width: 12, horizontalAlignment: XLAlignmentHorizontalValues.Center);
        FormatColumn(Heading.Future, width: 12, horizontalAlignment: XLAlignmentHorizontalValues.Center);
        FormatColumn(Heading.Step, width: 12, horizontalAlignment: XLAlignmentHorizontalValues.Center);
        FormatColumn(Heading.FirstRecordedTime, width: 22, horizontalAlignment: XLAlignmentHorizontalValues.Center);
        FormatColumn(Heading.FirstRecordedValue, width: 16);
        FormatColumn(Heading.CurrentTime, width: 22, horizontalAlignment: XLAlignmentHorizontalValues.Center);
        FormatColumn(Heading.CurrentValue, width: 16);
        FormatColumn(Heading.EngUnits, width: 12, horizontalAlignment: XLAlignmentHorizontalValues.Center);
        FormatColumn(Heading.CreationDate, width: 22, horizontalAlignment: XLAlignmentHorizontalValues.Center);

        // Special formatting adjusted for any DateTime column(s).
        FindColumnByHeading(Heading.CreationDate)?.Style.DateFormat.SetFormat(IsoTimeFormat);
        FindColumnByHeading(Heading.FirstRecordedTime)?.Style.DateFormat.SetFormat(IsoTimeFormat);
        FindColumnByHeading(Heading.CurrentTime)?.Style.DateFormat.SetFormat(IsoTimeFormat);

        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        IsSheetInitialized = true;
        IsDirty = true;
        Save();
    }

    private IXLColumn? FindColumnByHeading(string heading)
    {
        int index = FindColumnIndexByHeading(heading);
        return index == 0 ? null : DataSheet.Column(index);
    }

    private void FormatColumn(string heading, int? width = null, XLAlignmentHorizontalValues? horizontalAlignment = null)
    {
        var column = FindColumnByHeading(heading);
        if (column == null)
        {
            return;
        }

        if (width.HasValue)
        {
            if (width.Value <= 0)
            {
                column.Hide();
            }
            else
            {
                column.Width = width.Value;
            }
        }

        if (horizontalAlignment.HasValue)
        {
            column.Style.Alignment.Horizontal = horizontalAlignment.Value;
        }
    }

    private static void DrawBorders(IXLRange range)
    {
        var border = range.Style.Border;
        border.LeftBorder = XLBorderStyleValues.Thin;
        border.RightBorder = XLBorderStyleValues.Thin;
        border.TopBorder = XLBorderStyleValues.Thin;
        border.BottomBorder = XLBorderStyleValues.Thin;
        border.InsideBorder = XLBorderStyleValues.Thin;
    }

    public override void WriteRows(IDictionary<AFAttribute,
                                  PIPoint> map, PIPointList tags,
                                  IDictionary<PIPoint, AFValue> firstRecordedDict,
                                  IDictionary<PIPoint, AFValue> currentValueDict)
    {
        int previousCount = RowCount;

        // Between the AFAttribute and the PIPoint, all info has been gathered.
        // We are ready to publish a report row.
        foreach (var entry in map)
        {
            if (SkipRelativeAttribute(entry.Key))
            {
                continue;
            }
            AFValue? first = null;
            AFValue? current = null;
            if (entry.Value != null && Features.ShowFirstRecorded)
            {
                firstRecordedDict.TryGetValue(entry.Value, out first);
            }
            if (entry.Value != null && Features.ShowCurrentValue)
            {
                currentValueDict.TryGetValue(entry.Value, out current);
            }

            IncrementRowCount();
            IsDirty = true;
            int column = 1;
            foreach (var value in GetCellValues(entry.Key, entry.Value, first, current))
            {
                DataSheet.Cell(RowCount, column++).Value = XLCellValue.FromObject(value);
            }
        }

        if (RowCount > previousCount)
        {
            var range = DataSheet.Range(previousCount + 1, 1, RowCount, ColumnCount);
            DrawBorders(range);
        }
    }


    private int _saveCount = 0;

    public override void CustomSave()
    {
        // Set Filter on entire data range in sheet.
        var range = DataSheet.Range(1, 1, RowCount, ColumnCount);
        range.SetAutoFilter();

        if (_saveCount == 0)
        {
            Book.SaveAs(FileName);
        }
        else
        {
            Book.Save();
        }
        _saveCount++;
    }

    public override void CustomDispose()
    {
        Book.Dispose();
    }
}
