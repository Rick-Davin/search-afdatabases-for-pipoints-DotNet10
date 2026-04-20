namespace Search_AFDatabases_for_PIPoints.Logic;

internal class TextReport : ReportWriter
{
    private TextReport(AppFeatures features, DateTime creationTime, FileInfo outputFile) : base(features, creationTime)
    {
        this.OutputFile = outputFile;
    }

    public static TextReport CreateAndInitialize(AppFeatures features, DateTime? creationTime = default)
    {
        if (!creationTime.HasValue)
        {
            creationTime = DateTime.UtcNow;
        }
        var outputFile = new FileInfo($"SearchResults.{creationTime:yyyy-MM-dd}_{creationTime:HHmm}.tsv");
        var instance = new TextReport(features, creationTime.Value, outputFile);
        instance.Initialize();
        return instance;
    }

    private FileInfo OutputFile { get; }

    internal override string FileName => OutputFile.Name;
    internal override int MaximumRowCount => int.MaxValue;
    internal override int IndexStartsAt => 0;

    private static string Separator => "\t";

    private bool IsInitialized { get; set; }

    public override void Initialize()
    {
        if (IsInitialized)
        {
            return;
        }

        HeaderMap = CreateHeaderMap();

        File.WriteAllText(OutputFile.FullName, string.Join(Separator, HeaderMap.Keys) + Environment.NewLine);

        IsInitialized = true;
        IsDirty = false;
    }

    public override void WriteRows(IDictionary<AFAttribute,
                                  PIPoint> map, PIPointList tags,
                                  IDictionary<PIPoint, AFValue> firstRecordedDict,
                                  IDictionary<PIPoint, AFValue> currentValueDict)
    {
        var lines = new List<string>();

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
            var cellValues = new List<string>(capacity: ColumnCount);
            foreach (object? item in GetCellValues(entry.Key, entry.Value, first, current))
            {
                if (item == null)
                {
                    cellValues.Add(string.Empty);
                }
                else if (item is DateTime time)
                {
                    cellValues.Add(time.ToString(IsoTimeFormat));
                }
                else
                {
                    cellValues.Add(item.ToString() ?? "");
                }
            }

            var text = string.Join(Separator, cellValues);
            lines.Add(text);
        }

        if (lines.Count > 0)
        {
            // The file is already created so we only want to append to it.
            File.AppendAllLines(OutputFile.FullName, lines);
        }
    }
}
