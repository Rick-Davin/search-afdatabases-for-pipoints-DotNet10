namespace Search_AFDatabases_for_PIPoints.Logic
{
    internal class DatabaseWorker
    {
        public const string DataReferenceName = "PI Point";

        public DatabaseWorker(AFDatabase database, ReportWriter report, AppFeatures features, ILogger<MainWorker> logger)
        { 
            this.Database = database;
            this.Features = features;
            this.Report = report;
            this.Logger = logger;
        }

        private AFDatabase Database { get; }
        private ReportWriter Report { get; }
        private AppFeatures Features { get; }
        private ILogger<MainWorker> Logger { get; }
        private int TotalAttrCount { get; set; }

        public async Task Search()
        {
            long startTicks = Stopwatch.GetTimestamp();

            var page = new AFAttributeList();
            TotalAttrCount = 0;
            int pageIndex = 1;

            using (var search = new AFAttributeSearch(Database, null, $"Element:{{Name:'*'}} PlugIn:'{DataReferenceName}'"))
            {
                search.CacheTimeout = TimeSpan.FromMinutes(10);

                foreach (var attr in search.FindObjects(fullLoad: true, pageSize: Features.AFSearchPageSize))
                {
                    TotalAttrCount++;
                    page.Add(attr);
                    if (page.Count >= Features.AFSearchPageSize)
                    {
                        await ProcessPage(page, pageIndex++);
                        page = new AFAttributeList();
                    }
                }
            }

            if (page.Count > 0)
            {
                await ProcessPage(page, pageIndex);
            }

            Logger.LogInformation($"{Constant.Pad}End AFSearch for {TotalAttrCount} attributes.  Elapsed time = {Stopwatch.GetElapsedTime(startTicks)}");
        }

        private async Task ProcessPage(AFAttributeList page, int pageIndex)
        {
            long startTicks = Stopwatch.GetTimestamp();
            // Bulk data call.  Note some of the Value (PIPoint) may be null.
            var dict = page.GetPIPoint().Results;
            var tags = new PIPointList(dict.Values.Where(x => x != null));

            // Still gathering BULK info from RPC's to the PIServer.
            await tags.LoadAttributesAsync(Constant.PointAttributes);
            var firstRecordedDict = await GetFirstRecordedValueAsync(tags);
            var currentValueDict = await GetCurrentValueAsync(tags);

            Report.WriteRows(dict, tags, firstRecordedDict, currentValueDict);

            Logger.LogInformation($"{Constant.Pad}Page {pageIndex} has {page.Count} attributes.  Running Total = {TotalAttrCount}.  Page time = {Stopwatch.GetElapsedTime(startTicks)}");
        }

        private static async Task<IDictionary<PIPoint, AFValue>> GetFirstRecordedValueAsync(PIPointList tags)
        {
            var time = AFTime.MinValue;
            var values = (await tags.RecordedValueAsync(time, AFRetrievalMode.After)).Results;
            return ToDictionaryKeyedByPIPoint(values);
        }

        private static async Task<IDictionary<PIPoint, AFValue>> GetCurrentValueAsync(PIPointList tags)
        {
            IDictionary<PIPoint, AFValue> dict = new Dictionary<PIPoint, AFValue>(capacity: tags.Count);

            // So PIPointList does not have a CurrentValueAsync method.
            // We will make 2 calls, one async for historical tags and one blocking for future tags.
            var futureTags = new PIPointList(tags.Where(x => x.Future));
            var historyTags = new PIPointList(tags.Where(x => !x.Future));

            if (historyTags.Count > 0)
            {
                var historyValues = (await tags.EndOfStreamAsync()).Results;
                dict = ToDictionaryKeyedByPIPoint(historyValues);
            }

            if (futureTags.Count > 0)
            {
                var futureValues = futureTags.CurrentValue();
                foreach (var value in futureValues)
                {
                    // Avoid using Add() here in case of duplicate PIPoints.
                    dict[value.PIPoint] = value;
                }
            }

            return dict;
        }

        private static IDictionary<PIPoint, AFValue> ToDictionaryKeyedByPIPoint(IList<AFValue> values)
        {
            var dict = new Dictionary<PIPoint, AFValue>(capacity: values.Count);
            foreach (var value in values)
            {
                // It is possible to reference same PIPoint on different AFAttributes,
                // which could produce duplicate entries.  Therefore, I do not use the
                // Add method, and instead use direct assignment, so the entry will 
                // be overridden.
                dict[value.PIPoint] = value;
            }
            return dict;
        }

    }
}
