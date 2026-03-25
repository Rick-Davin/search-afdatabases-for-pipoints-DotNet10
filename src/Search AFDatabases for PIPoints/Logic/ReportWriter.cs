using Search_AFDatabases_for_PIPoints.Models;
using OSIsoft.AF.Asset;
using OSIsoft.AF.PI;
using System.Timers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OSIsoft.AF.Time;

namespace Search_AFDatabases_for_PIPoints.Logic
{
    internal abstract class ReportWriter : IReportWriter, IDisposable
    {
        private protected ReportWriter(AppFeatures features, DateTime creationTime)
        {
            this.Features = features;
            this.CreationTime = creationTime;
            this.UseAutoSave = features.UseReportAutoSave;
            if (features.UseReportAutoSave)
            {
                this.AutoSaveDuration = TimeSpan.FromSeconds(features.AutoSaveSeconds);
            }
            RowCount = 1;
        }

        internal virtual AppFeatures Features { get; }

        internal virtual DateTime CreationTime { get; }
        internal abstract string FileName { get; }

        internal virtual IDictionary<string, int>? HeaderMap { get; set; }

        internal abstract int MaximumRowCount { get; }
        internal virtual int IndexStartsAt { get; }
        internal virtual int RowCount { get; set; }
        internal virtual int ColumnCount => HeaderMap?.Count ?? 0;

        internal bool IsDirty { get; set; }

        public const string IsoTimeFormat = "yyyy-mm-dd hh:mm:ss";

        internal bool UseAutoSave { get; }
        internal TimeSpan AutoSaveDuration { get; }

        internal System.Timers.Timer? SaveTrigger { get; set; } = null;

        public abstract void WriteRows(IDictionary<AFAttribute, PIPoint> map, PIPointList tags, IDictionary<PIPoint, AFValue> firstRecordedDict, IDictionary<PIPoint, AFValue> currentValueDict);

        public abstract void Initialize();

        internal static class Heading
        {
            // AFAttribute-related fields
            public const string AssetServer = "Asset Server";
            public const string Database = "AF Database";
            public const string ElementPath = "Element Path";
            public const string AttributePath = "Attribute Path";
            public const string AttributeGuid = "Attribute Guid";
            public const string AttributeDescription = "Description";
            public const string AttributeType = "Data Type";
            public const string DefaultUom = "Default UOM";
            public const string SourceUom = "Source UOM";
            public const string DataReference = "Data Reference";
            public const string ConfigString = "Config String";
            // PIPoint-related fields
            public const string DataArchive = "Data Archive";
            public const string Tag = "Tag";
            public const string PointId = "PointID";
            public const string Future = "Future";
            public const string Step = "Step";
            public const string TagDescriptor = "Descriptor";
            public const string PointType = "PointType";
            public const string CreationDate = "Creation Date";
            public const string EngUnits = "EngUnits";
            // OPTIONAL PIPoint-related fields
            public const string TagGrouping = "Grouping";
            public const string FirstRecordedTime = "First Recorded Time";
            public const string FirstRecordedValue = "First Recorded Value";
            public const string CurrentTime = "Current Time";
            public const string CurrentValue = "Current Value";
        } // end Heading subclass

        // The order that appears here is the order the columns will appear on the Excel Report
        // AFTER any AF-related Headings.
        internal virtual HashSet<string> CreateOrderedPIHeadings()
        {
            var set = new HashSet<string>(capacity: 20, StringComparer.OrdinalIgnoreCase)
                { Heading.DataArchive };

            if (Features.TagGroupingSeparators!.Count > 0)
            {
                set.Add(Heading.TagGrouping);
            }

            foreach (string heading in new List<string>()
                     {
                         Heading.DataArchive,
                         Heading.TagGrouping,
                         Heading.Tag,
                         Heading.TagDescriptor,
                         Heading.PointId,
                         Heading.Future,
                         Heading.Step,
                         Heading.PointType
                     })
            {
                set.Add(heading);
            }

            if (Features.ShowFirstRecorded)
            {
                set.Add(Heading.FirstRecordedTime);
                set.Add(Heading.FirstRecordedValue);
            }
            if (Features.ShowCurrentValue)
            {
                set.Add(Heading.CurrentTime);
                set.Add(Heading.CurrentValue);
            }
            set.Add(Heading.EngUnits);
            set.Add(Heading.CreationDate);
            return set;
        }

        internal virtual IDictionary<string, int> CreateHeaderMap()
        {
            if (OrderedPIHeadings == null)
            {
                OrderedPIHeadings = CreateOrderedPIHeadings();
            }

            int capacity = OrderedAFHeadings.Count + OrderedPIHeadings.Count;

            // Case-insensitive dictionary.  The Value will be the Data Sheet's repective Column Index for each Heading.
            var map = new Dictionary<string, int>(capacity: capacity, comparer: StringComparer.OrdinalIgnoreCase);
            int i = IndexStartsAt;  // Excel is 1-based, Text is 0-based
            foreach (string heading in OrderedAFHeadings)
            {
                map.Add(heading, i++);
            }
            foreach (string heading in OrderedPIHeadings)
            {
                map.Add(heading, i++);
            }

            return map;
        }

        // The order that appears here is the order the columns will appear on the Excel Report.
        internal static HashSet<string> OrderedAFHeadings => new HashSet<string>(comparer: StringComparer.OrdinalIgnoreCase)
            {
                Heading.AssetServer,
                Heading.Database,
                Heading.ElementPath,
                Heading.AttributePath,
                Heading.AttributeGuid,
                Heading.AttributeDescription,
                Heading.AttributeType,
                Heading.DefaultUom,
                Heading.SourceUom,
                Heading.DataReference,
                Heading.ConfigString
            };

        // The order that appears here is the order the columns will appear on the Excel Report
        // AFTER any AF-related Headings.
        internal HashSet<string>? OrderedPIHeadings { get; set; }



        internal int FindColumnIndexByHeading(string heading)
        {
            HeaderMap!.TryGetValue(heading, out int index);
            return index;
        }

        internal void IncrementRowCount()
        {
            if (RowCount == MaximumRowCount)
            {
                Save(disposing: true);
                throw new ArgumentOutOfRangeException($"The Excel Report is limited to {MaximumRowCount} rows.  Unable to continue.");
            }
            RowCount++;
        }

        internal bool SkipRelativeAttribute(AFAttribute attr)
        {
            if (Features.ShowRelativePIPoints)
            {
                return false;
            }
            return !attr.ConfigString.StartsWith(@"\\");
        }


        internal IEnumerable<object?> GetAttributeCellValues(AFAttribute attr)
        {
            foreach (var entry in HeaderMap!)
            {
                switch (entry.Key)
                {
                    case Heading.AssetServer:
                        yield return attr.PISystem.Name;
                        break;
                    case Heading.Database:
                        yield return attr.Database.Name;
                        break;
                    case Heading.ElementPath:
                        yield return attr.Element.GetPath(attr.Database);
                        break;
                    case Heading.AttributePath:
                        yield return attr.GetPath(attr.Element);
                        break;
                    case Heading.AttributeGuid:
                        yield return attr.ID.ToString();
                        break;
                    case Heading.AttributeDescription:
                        yield return attr.Description;
                        break;
                    case Heading.AttributeType:
                        yield return attr.Type.Name;
                        break;
                    case Heading.DefaultUom:
                        yield return attr.DefaultUOM?.Abbreviation;
                        break;
                    case Heading.SourceUom:
                        yield return attr.SourceUOM?.Abbreviation;
                        break;
                    case Heading.DataReference:
                        yield return attr.DataReferencePlugIn?.Name;
                        break;
                    case Heading.ConfigString:
                        yield return attr.ConfigString;
                        break;
                }
            }
        }

        internal IEnumerable<object?> GetTagCellValues(PIPoint tag, AFValue? firstRecorded, AFValue? current)
        {
            var dict = tag.GetAttributes(Constant.PointAttributes);
            foreach (var entry in HeaderMap!)
            {
                switch (entry.Key)
                {
                    case Heading.DataArchive:
                        yield return tag.Server.Name;
                        break;
                    case Heading.PointId:
                        yield return tag.ID;
                        break;
                    case Heading.Future:
                        yield return tag.Future;
                        break;
                    case Heading.Step:
                        yield return tag.Step;
                        break;
                    case Heading.TagGrouping:
                        yield return GetGrouping(tag);
                        break;
                    case Heading.Tag:
                        yield return tag.Name;
                        break;
                    case Heading.TagDescriptor:
                        yield return (string)dict[PICommonPointAttributes.Descriptor];
                        break;
                    case Heading.PointType:
                        yield return tag.PointType.ToString();
                        break;
                    case Heading.EngUnits:
                        yield return (string)dict[PICommonPointAttributes.EngineeringUnits];
                        break;
                    case Heading.CreationDate:
                        yield return (DateTime)dict[PICommonPointAttributes.CreationDate];
                        break;
                    case Heading.FirstRecordedTime:
                        yield return firstRecorded?.Timestamp.UtcTime;
                        break;
                    case Heading.FirstRecordedValue:
                        yield return ToCleanerObject(firstRecorded);
                        break;
                    case Heading.CurrentTime:
                        yield return current?.Timestamp.UtcTime;
                        break;
                    case Heading.CurrentValue:
                        yield return ToCleanerObject(current);
                        break;
                }
            }
        }

        internal IEnumerable<object?> GetCellValues(AFAttribute attr, PIPoint? tag, AFValue? first, AFValue? current)
        {
            foreach (object? item in GetAttributeCellValues(attr))
            {
                yield return item;
            }

            if (tag != null)
            {
                foreach (object? item in GetTagCellValues(tag, first, current))
                {
                    yield return item;
                }
            }
        }

        internal static object? ToCleanerObject(AFValue? value)
        {
            if (value == null)
            {
                return null;
            }
            var digitalState = value.Value as AFEnumerationValue;
            if (digitalState != null)
            {
                return digitalState.Name;
            }
            if (value.Value is AFTime aftime)
            {
                return aftime.UtcTime;
            }
            if (value.Value is DateTime time)
            {
                return time.ToUniversalTime();
            }
            // I said cleaner, not absolutely cleaned.
            return value.Value;
        }

        internal string GetGrouping(PIPoint tag)
        {
            // Grouping is a custom way to help filter on tens of thousands of tags.
            // The TagGroupingSeparators is a configurable setting to hopefully match your
            // own company's tag naming and multiple separators can help with naming standards
            // that have changed over time.  Consider an example where some tags begin
            // with thousands of tags beginning each with
            //      "Site 1-*"
            //      "Site 2-*"
            //      "Site1/*"
            //      "Site2/*"
            // Using separators of "-" and "/", the tags will be grouped respectively as
            //      "Site 1"
            //      "Site 2"
            //      "Site1"
            //      "Site2"
            // Which can be immensely helpful using Excel filtering.

            if (Features.TagGroupingSeparators == null || Features.TagGroupingSeparators.Count == 0)
            {
                return string.Empty;
            }

            int length = 0;
            foreach (var separator in Features.TagGroupingSeparators)
            {
                int index = tag.Name.IndexOf(separator);
                if (index > 0 && (index < length || length <= 0))
                {
                    length = index;
                }
            }

            return length > 0 ? tag.Name.Substring(0, length) : string.Empty;
        }

        private bool _disposedValue;

        private void SetSaveTrigger(bool activate)
        {
            if (!activate || !UseAutoSave)
            {
                SaveTrigger = null;
            }
            // Create a timer with a two second interval.
            SaveTrigger = new System.Timers.Timer(AutoSaveDuration.TotalMilliseconds);
            if (SaveTrigger != null)
            {
                // Hook up the Elapsed event for the timer. 
                SaveTrigger.Elapsed += AutoSaveTriggered;
                SaveTrigger.AutoReset = false;
                SaveTrigger.Enabled = true;
            }
        }

        private void AutoSaveTriggered(Object? source, ElapsedEventArgs e)
        {
            SetSaveTrigger(false);
            Save();
        }


        private readonly object _lockObject = new object();
        public void Save(bool disposing = false)
        {
            lock (_lockObject)
            {
                if (!IsDirty)
                {
                    return;
                }

                SetSaveTrigger(false);

                CustomSave();

                IsDirty = false;
                SetSaveTrigger(!disposing);
            }
        }

        public virtual void CustomSave() { /*nop*/ }

        #region Disposable
        public virtual void CustomDispose() { /*nop*/ }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposedValue)
            {
                if (disposing)
                {
                    // TODO: dispose managed state (managed objects)
                    Save(disposing);
                    CustomDispose();
                }

                // TODO: free unmanaged resources (unmanaged objects) and override finalizer
                // TODO: set large fields to null
                _disposedValue = true;
            }
        }

        public void Dispose()
        {
            // Do not change this code. Put cleanup code in 'Dispose(bool disposing)' method
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }



        #endregion

    }
}
