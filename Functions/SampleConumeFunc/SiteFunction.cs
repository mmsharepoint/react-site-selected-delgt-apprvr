using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using SampleConsumeFunc.Services;

namespace SampleConsumeFunc
{
  public class SiteFunction
  {
    private readonly ITokenValidationService _tokenValidationService;
    private readonly IGraphService _graphClientService;
    private readonly ILogger<SiteFunction> _logger;
    private readonly IConfiguration _config;

    public SiteFunction(ITokenValidationService tokenValidationService, IGraphService graphClientService, ILogger<SiteFunction> logger, IConfiguration config)
    {
      _tokenValidationService = tokenValidationService;
      _graphClientService = graphClientService;
      _logger = logger;
      _config = config;
    }

    [Function("SiteFunction")]
    public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
    {
      _logger.LogInformation("C# HTTP trigger function processed a request.");

      var bearerToken = await _tokenValidationService
          .ValidateAuthorizationHeaderAsync(req);
      _logger.LogInformation("Bootstrap token: " + bearerToken); // not nessesary

      string siteUrl = req.Query["URL"];
      bool siteDescreption = await _graphClientService.UpdateSiteDescreption(bearerToken, siteUrl);
      return new OkObjectResult("Welcome to Azure Functions!");
    }
  }
}
