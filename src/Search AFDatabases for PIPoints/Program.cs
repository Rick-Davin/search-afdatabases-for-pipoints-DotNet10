namespace Search_AFDatabases_for_PIPoints
{
    // What is going on here?  This does not look like my old AF SDK console application in .NET Framework.
    //
    // We are not using .NET Framework anymore, so we no longer will have an AppConfig file at our disposal.
    // Instead, we will use appsettings.json for our configuration settings, and also use the preferred ILogger<T> pattern
    // for logging to the console as well as to a file. With all that going on, we will also use the Generic Host
    // and Dependency Injection (DI) patterns to structure our application.
    //
    // Is this neccessary for thie sample application?
    //
    // No, but it is a good practice to follow for modern .NET applications, and it will make it easier to scale and maintain our application as it grows in complexity.
    // For instance, this application or another application in the future may require appsettings.{ENVIRONMENT}.json files, command line arguments, environment variables,
    // or a secrets file.  By using the Generic Host and DI patterns, we can easily add these configuration sources and manage our application's dependencies in a clean and maintainable way.  

    internal class Program  
    {
        static async Task Main(string[] args)
        {
            Environment.CurrentDirectory = AppDomain.CurrentDomain.BaseDirectory;

            using IHost host = Host.CreateDefaultBuilder(args)
                .ConfigureAppConfiguration((hostingContext, config) =>
                {
                    config.SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                          .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
                          .AddEnvironmentVariables();
                })
                .ConfigureServices((context, services) =>
                {
                    services.Configure<AppFeatures>(context.Configuration.GetSection("Features"));
                    // Register application services
                    services.AddTransient<MainWorker>();
                })
                .UseSerilog((context, loggerConfig) =>
                {
                    string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                    string logPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs", $"app-{timestamp}.txt");
                    
                    loggerConfig
                        .MinimumLevel.Information()
                        .WriteTo.Console(
                            outputTemplate: "[{Timestamp:HH:mm:ss.fff}] [{Level:u3}] {Message:lj}{NewLine}{Exception}")
                        .WriteTo.File(
                            path: logPath,
                            outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss.fffff}] [{Level:u3}] {Message:lj}{NewLine}{Exception}");
                })
                .Build();

            // Create a scope and run the AppWorker resolved from DI.
            // The IConfiguration (from appsettings.json) and ILogger will be provided to AppWorker by the DI container.
            using var scope = host.Services.CreateScope();
            var provider = scope.ServiceProvider;

            // What would have been the Main() method in a .NET Framework application is now the DoWorkAsync() method in the MainWorker class.  
            MainWorker worker = provider.GetRequiredService<MainWorker>();
            await worker.DoWorkAsync();

            Console.WriteLine("\n\nPress ENTER key to exit...");
            Console.ReadLine();
        }
    }
}
