using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Services;
using Utilities;

namespace ConsoleApplication;

internal static partial class Program
{
    private static void Main()
    {
        using IHost host = CreateHostBuilder().Build();
        using var scope = host.Services.CreateScope();

        var services = scope.ServiceProvider;
        var logger = services.GetRequiredService<ILogger>();
        try
        {
            services.GetRequiredService<Application>().Run();
        }
        catch (Exception ex)
        {
            logger.LogErrorAndExit(ex.ToString());
        }
    }

    private static IHostBuilder CreateHostBuilder()
    {
        return Host.CreateDefaultBuilder()
            .ConfigureServices((_, services) =>
            {
                services.AddSingleton<ILogger, ConsoleLogger>();
                services.AddSingleton<DataService>();
                services.AddSingleton<ReadService>();
                services.AddSingleton<WriteService>();
                services.AddSingleton<ExportService>();
                services.AddSingleton<Application>();
            });
    }
}