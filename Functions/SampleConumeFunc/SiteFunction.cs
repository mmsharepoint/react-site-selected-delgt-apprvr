using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace SampleConsumeFunc
{
  public class SiteFunction
  {
    private readonly ILogger<SiteFunction> _logger;
    private readonly IConfiguration _config;

    public SiteFunction(ILogger<SiteFunction> logger, IConfiguration config)
    {
      _logger = logger;
      _config = config;
    }

    [Function("Function1")]
    public IActionResult Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequest req)
    {
      _logger.LogInformation("C# HTTP trigger function processed a request.");
      return new OkObjectResult("Welcome to Azure Functions!");
    }
  }
}
