using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using OSIsoft.AF;
using Search_AFDatabases_for_PIPoints.Models;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

// Why this extra layer for MainWorker?  
// MainWorker serves as the orchestrator of the entire application workflow, and is responsible for coordinating the various components and steps of the application.
// By that definition, MainWork could be used in Program.Main, but even if I was using .NET Framework, I would keep MainWorker as a separate class and keep
// Program.Main as minimal as possible.
//
// By having an MainWorker class, we can encapsulate all the logic related to orchestrating the workflow of the application in one place.
// This makes it easier to maintain and test the code, as well as to add additional functionality in the future without cluttering the Main method.

namespace Search_AFDatabases_for_PIPoints.Logic
{
    // The purpose of the MainWorker is to:
    //      (1) Validate features from appsettings.json
    //      (2) Display start-up information
    //      (3) Connect to the Asset Server (PISystem)
    //      (4) Initialize the ReportWriter (Excel or Text)
    //      (5) Find the AFDatabase(s) to search based on the requested search filter
    //      (6) Finally, loop through each found AFDatabase and calling DatabaseWorker with each one.
    internal class MainWorker
    {
        private AppFeatures Features { get; }
        private ILogger<MainWorker> Logger { get; }

        public MainWorker(IOptions<AppFeatures> options, ILogger<MainWorker> logger)
        {
            // While the features are read from the appsettings.json, we access them via the IOptions<Features> interface (see Program.Main) in the name of Dependency Injection.
            Features = ValidateFeatures(options.Value);
            Logger = logger;
        }

        #region Public Methods
        public async Task DoWorkAsync()
        {
            long startTicks = Stopwatch.GetTimestamp();

            try
            {
                DisplayPreamble();

                PISystem? assetServer = FindAssetServer(Features.AssetServer);
                if (assetServer == null)
                {
                    return;
                }

                List<string>? databaseNames = FindMatchingDatabaseNames(assetServer, Features.DatabaseSearchPattern);

                // The output directory will first be treated as relative to the current directory, but you can specficy a fully qualified path if you want the output to be written somewhere else.  
                // That is why it is called Path instead of Name. 
                Environment.CurrentDirectory = AppDomain.CurrentDomain.BaseDirectory;
                DirectoryInfo outputFolder = new DirectoryInfo(Features.OutputFolderPath);
                if (!outputFolder.Exists)
                {
                    outputFolder.Create();
                }


                ReportWriter? report = InitializeReport(Features, databaseNames.Count, outputFolder);

                if (report != null)
                {
                    for (int i = 0; i < databaseNames.Count; i++)
                    {
                        AFDatabase database = assetServer.Databases[databaseNames[i]];
                        Logger.LogInformation(Constant.DashLine);
                        Logger.LogInformation("BEGIN AFDATABASE ({index}) - {name}", i + 1, database.Name);
                        var worker = new DatabaseWorker(database, report, Features, Logger);
                        await worker.Search();

                        if (Logger.IsEnabled(LogLevel.Information))
                        {
                            var elapsed = Stopwatch.GetElapsedTime(startTicks);
                            Logger.LogInformation("END AFDATABASE ({index})  Running time = {elapsed}", i + 1, elapsed);
                        }
                    }

                    long mark = Stopwatch.GetTimestamp();
                    Logger.LogInformation(Constant.DashLine);
                    Logger.LogInformation("Saving final report to disk ...");
                    report.Save();
                    if (Logger.IsEnabled(LogLevel.Information))
                    {
                        var saveElapsed = Stopwatch.GetElapsedTime(mark);
                        Logger.LogInformation("{pad}save took {elapsed}", Constant.Pad, saveElapsed);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.LogError(ex, "An unhandled exception was caught in DoWorkAsync.");
            }

            Logger.LogInformation("\n{dashes}\nEND OF APPLICATION.  OVERALL ELAPSED TIME = {elapsed}\n{dashes}", Constant.DashLine, Stopwatch.GetElapsedTime(startTicks), Constant.DashLine);
        }

        #endregion

        #region Private Methods

        private static AppFeatures ValidateFeatures(AppFeatures features)
        {
            features.TagGroupingSeparators = GetValidSeparators(features.TagGroupingSeparators);

            if (features.UseReportAutoSave && features.AutoSaveSeconds <= 0)
            {
                const int DefaultAutoSeconds = 120;
                features.AutoSaveSeconds = DefaultAutoSeconds;
            }

            if ("Excel".Equals(features.OutputWriter ?? "", StringComparison.OrdinalIgnoreCase))
            {
                features.OutputWriter = "Excel";
            }
            else
            {
                features.OutputWriter = "Text";
            }

            return features;
        }

        private static string GetFrameworkDescription()
        {
            return System
                .Runtime
                .InteropServices
                .RuntimeInformation
                .FrameworkDescription;
        }

        private static List<string> GetValidSeparators(List<string>? separators)
        {
            if (separators == null)
            {
                return new List<string>();
            }
            // Quietly removed any empty strings.  NOTE: white spaces are allowed as separators so do not use IsNullEmptyOrWhitespace.
            for (int i = separators.Count - 1; i >= 0; i--)
            {
                if (string.IsNullOrEmpty(separators[i]))
                {
                    separators.RemoveAt(i);
                }
            }
            return new List<string>(separators);
        }


        private void DisplayPreamble()
        {
            if (!Logger.IsEnabled(LogLevel.Information))
            {
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine("APP STARTING ... \nThis application search for all AFAttributes using the PI Point data reference");
            sb.AppendLine("for one or more AFDatabases on the same Asset Server (PISystem).");
            sb.AppendLine("Found information will be written to an Excel workbook.");
            sb.AppendLine(Constant.DashLine);
            sb.AppendLine("INPUT CONFIGURATION found in appsettings.json");
            sb.AppendLine($"{Constant.Pad}Asset Server Name = \"{Features.AssetServer}\"");
            sb.AppendLine($"{Constant.Pad}Database Search Pattern = \"{Features.DatabaseSearchPattern}\"");
            sb.AppendLine($"{Constant.Pad}Output Folder Path = \"{Features.OutputFolderPath}\"");
            sb.AppendLine($"{Constant.Pad}AF Search Page Size = {Features.AFSearchPageSize}");
            sb.AppendLine($"{Constant.Pad}Show Relative PIPoints = {Features.ShowRelativePIPoints}");
            sb.AppendLine($"{Constant.Pad}Show Current Values = {Features.ShowCurrentValue}");

            sb.AppendLine($"{Constant.Pad}Show Show First Recorded Values = {Features.ShowFirstRecorded}");
            if (Features.ShowFirstRecorded)
            {
                // The previous AppendLine has the carriage return + line feed
                Logger.LogInformation(sb.ToString());
                Logger.LogInformation("{pad}*** WARNING: this setting may degrade performance. ***", Constant.Pad);
                sb.Clear();
            }

            sb.AppendLine($"{Constant.Pad}Output Report Writer = {Features.OutputWriter}");
            sb.AppendLine($"{Constant.Pad}Use Report AutoSave = {Features.UseReportAutoSave}");
            if (Features.UseReportAutoSave)
            {
                sb.AppendLine($"{Constant.Pad}Auto Save Report every = {Features.AutoSaveSeconds} seconds");
            }

            sb.AppendLine($"{Constant.Pad}Tag Grouping Separators:");
            if (Features.TagGroupingSeparators!.Count == 0)
            {
                sb.AppendLine($"{Constant.Pad}{Constant.Pad}<none>");
            }
            else
            {
                foreach (var item in Features.TagGroupingSeparators)
                {
                    sb.AppendLine($"{Constant.Pad}{Constant.Pad}\"{item}\"");
                }
            }
            sb.AppendLine(Constant.DashLine);
            sb.AppendLine($"Is Host 64-bit Operating System: {Environment.Is64BitOperatingSystem}");
            sb.AppendLine($"Is this App a 64-bit Process: {Environment.Is64BitProcess}");
            var afclient = new PISystems();
            sb.AppendLine($"Host AF SDK Version = {afclient.Version}");
            sb.AppendLine($".NET Description: {GetFrameworkDescription()}");
            sb.AppendLine($"Common Language Runtime Version = {Environment.Version}");
            // The previous AppendLine has the carriage return + line feed
            Logger.LogInformation(sb.ToString());
        }

        private PISystem? FindAssetServer(string name)
        {
            long startTicks = Stopwatch.GetTimestamp();

            Logger.LogInformation(Constant.DashLine);
            Logger.LogInformation("Connecting to Asset Server '{name}' ...", name);
            PISystem? assetServer = AFOperation.ConnectToAssetServer(name);
            if (assetServer == null)
            {
                Logger.LogCritical("Unable to connect to Asset Server");
                return null;

            }

            if (Logger.IsEnabled(LogLevel.Information))
            {
                var elapsed = Stopwatch.GetElapsedTime(startTicks);
                Logger.LogInformation("{pad}connection took {seconds} seconds", Constant.Pad, elapsed.TotalSeconds);
            }

            var sb = new StringBuilder();
            sb.AppendLine($"AF Server Version: {assetServer.ServerVersion}");
            sb.AppendLine($"Required Version : 2.9.5 or later");

            // The previous AppendLine has the carriage return + line feed
            Logger.LogInformation(sb.ToString());

            string errorMsg = string.Empty;

            if (!assetServer.Supports(PISystemFeatures.QuerySearchAttribute))
            {
                errorMsg = $"ERROR: Your Asset Server version does not support the AFAttributeSearch method.";
            }
            else if (!IsMinimumVersion(assetServer.ServerVersion))
            {
                errorMsg = $"ERROR: While your Asset Server version does support the AFAttributeSearch method," +
                         "\nit does not support the PlugIn filter.";
            }

            if (!string.IsNullOrWhiteSpace(errorMsg))
            {
                Logger.LogCritical(errorMsg);
                return null;
            }

            return assetServer;
        }

        private List<string> FindMatchingDatabaseNames(PISystem assetServer, string pattern)
        {
            long startTicks = Stopwatch.GetTimestamp();
            Logger.LogInformation(Constant.DashLine);
            Logger.LogInformation("Searching for AFDatabases matching name pattern '{pattern}' ...", pattern);
            var list = new List<string>(capacity: 10);

            int count = 0;
            foreach (string name in assetServer.GetMatchingDatabaseByNameOrPattern(pattern).Select(x => x.Name))
            {
                count++;
                Logger.LogInformation("{pad}({index}) {dbName}", Constant.Pad, count, name);
                list.Add(name);
            }

            if (count == 0)
            {
                Logger.LogWarning("{pad}No databases were found that matched the name pattern!", Constant.Pad);
            }

            if (Logger.IsEnabled(LogLevel.Information))
            {
                var elapsed = Stopwatch.GetElapsedTime(startTicks);
                Logger.LogInformation("Finding matching databases took {seconds} seconds", elapsed.TotalSeconds);
            }

            return list;
        }

        private static bool IsMinimumVersion(string version)
        {
            // For using PlugIn filter on the AFAttributeSearch method, the minimum AF version is 2.9.5.
            var tokens = version.Split('.');
            int major = int.Parse(tokens[0]);
            if (major < 2)
            {
                return false;
            }
            int minor = int.Parse(tokens[1]);
            if (minor < 9)
            {
                return false;
            }
            if (minor > 9)
            {
                return true;
            }
            int revision = tokens.Length < 3 ? 0 : int.Parse(tokens[2]);
            return revision >= 5;
        }

        private ReportWriter? InitializeReport(AppFeatures features, int databaseCount, DirectoryInfo outputFolder)
        {
            Directory.SetCurrentDirectory(outputFolder.FullName);
            ReportWriter? report = null;
            Logger.LogInformation(Constant.DashLine);
            Logger.LogInformation("Initializing {writer} Report Writer ...", features.OutputWriter);
            if (databaseCount == 0)
            {
                Logger.LogWarning("{pad}Report not created since 0 databases were found.", Constant.Pad);
                return report;
            }

            long startTicks = Stopwatch.GetTimestamp();
            if ("Excel".Equals(features.OutputWriter, StringComparison.OrdinalIgnoreCase))
            {
                report = ExcelReport.CreateAndInitialize(features);
            }
            else
            {
                report = TextReport.CreateAndInitialize(features);
            }

            if (Logger.IsEnabled(LogLevel.Information))
            {
                var elapsed = Stopwatch.GetElapsedTime(startTicks);
                Logger.LogInformation("{pad}Report initialization took {seconds} seconds", Constant.Pad, elapsed.TotalSeconds);
            }

            return report;
        }
        #endregion

    }
}
