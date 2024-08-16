using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using ProjectReportConsoleApp;
using Services;
using Utilities;

using IHost host = CreateHostBuilder().Build();
using var scope = host.Services.CreateScope();

var services = scope.ServiceProvider;
var logger = services.GetRequiredService<ILogger>();
var application = services.GetRequiredService<Application>();
try
{
    application.Run();
}
catch (Exception ex)
{
    logger.LogError(ex.ToString());
    application.ExitApplication();
}

static IHostBuilder CreateHostBuilder()
{
    return Host
        .CreateDefaultBuilder()
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