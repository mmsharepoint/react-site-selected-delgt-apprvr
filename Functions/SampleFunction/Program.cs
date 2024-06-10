using System.Reflection;
using System.Text.Json;
using Azure.Core.Serialization;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using SampleFunction.Services;

var host = new HostBuilder()
  .ConfigureFunctionsWorkerDefaults(configureOptions: options =>
  {
    options.Serializer = new JsonObjectSerializer(
      new JsonSerializerOptions
      {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
      }
    );
  })
  .ConfigureAppConfiguration(config =>
  {
    config.AddUserSecrets(Assembly.GetExecutingAssembly(), false);
  })
  .ConfigureServices(services => {
    //services.AddSingleton<IGraphService, GraphService>();
    //services.AddSingleton<ITokenValidationService, TokenValidationService>();
  })
  .Build();

host.Run();