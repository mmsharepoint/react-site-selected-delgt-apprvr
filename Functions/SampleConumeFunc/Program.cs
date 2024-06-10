using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var host = new HostBuilder()
  .ConfigureFunctionsWebApplication()
  .ConfigureAppConfiguration((hostContext, config) =>
  {
    config.AddJsonFile("host.json", optional: true);
  })
  .ConfigureServices(services =>
  {
    services.AddApplicationInsightsTelemetryWorkerService();
    services.ConfigureFunctionsApplicationInsights();
  })
  .Build();

host.Run();
