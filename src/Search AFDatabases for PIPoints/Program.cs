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

using var scope = host.Services.CreateScope();
var provider = scope.ServiceProvider;

MainWorker worker = provider.GetRequiredService<MainWorker>();
await worker.DoWorkAsync();

Console.WriteLine("\n\nPress ENTER key to exit...");
Console.ReadLine();
