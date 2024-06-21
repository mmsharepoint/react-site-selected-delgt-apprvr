using Azure.Core;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using ApplyPermissions.Services;
using ApplyPermissions.Model;
using System.Text;
using System.Text.Json;

namespace ApplyPermissions
{
  public class ApplyPermimssion
  {
    private readonly ITokenValidationService _tokenValidationService;
    private readonly IGraphService _graphClientService;
    private readonly ILogger<ApplyPermimssion> _logger;
    private readonly IConfiguration _config;

    public ApplyPermimssion(ITokenValidationService tokenValidationService, IGraphService graphClientService, ILogger<ApplyPermimssion> logger, IConfiguration config)
    {
      _tokenValidationService = tokenValidationService;
      _graphClientService = graphClientService;
      _logger = logger;
      _config = config;
    }

    [Function("ApplyPermimssion")]
    public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
    {
      _logger.LogInformation("C# HTTP trigger function processed a request.");
      var bearerToken = await _tokenValidationService
          .ValidateAuthorizationHeaderAsync(req);

      _logger.LogInformation("Bootstrap token: " + bearerToken);

      string bodyContents;
      using (Stream receiveStream = req.Body)
      {
        using (StreamReader readStream = new StreamReader(receiveStream, Encoding.UTF8))
        {
          bodyContents = readStream.ReadToEndAsync().Result;
        }
      }
      var body = JsonSerializer.Deserialize<Model.Request>(bodyContents);
      string siteUrl = body.URL;
      string role = body.Permission;
      string appId = body.AppID;

      bool siteDescreptionUpdated = await _graphClientService.ApplySitePermission(bearerToken, siteUrl, role, appId);

      return new OkObjectResult($"Welcome to Azure Functions!");
    }

    [Function("GetServicePrincipals")]
    public async Task<IActionResult> GetServicePrincipals([HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequestData req)
    {
      _logger.LogInformation("C# HTTP trigger function processed a request.");
      var bearerToken = await _tokenValidationService
          .ValidateAuthorizationHeaderAsync(req);

      _logger.LogInformation("Bootstrap token: " + bearerToken);

      string appPrefix = req.Query["Prefix"];

      var appRegs = await _graphClientService.GetServicePrincipals(bearerToken, appPrefix);

      return new OkObjectResult(appRegs);
    }

    [Function("SearchSites")]
    public async Task<IActionResult> SearchSites([HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequestData req)
    {
      _logger.LogInformation("C# HTTP trigger function processed a request.");
      var bearerToken = await _tokenValidationService
          .ValidateAuthorizationHeaderAsync(req);

      string queryText = req.Query["QueryText"];

      var appRegs = await _graphClientService.SearchSites(bearerToken, queryText);

      return new OkObjectResult(appRegs);
    }

    [Function("IsSiteAdmin")]
    public async Task<IActionResult> IsSiteAdmin([HttpTrigger(AuthorizationLevel.Anonymous, "get")] HttpRequestData req)
    {
      _logger.LogInformation("C# HTTP trigger function processed a request.");
      var bearerToken = await _tokenValidationService
          .ValidateAuthorizationHeaderAsync(req);

      string siteId = req.Query["SiteId"];
      string userMail = req.Query["User"];

      var appRegs = await _graphClientService.IsSiteAdmin(bearerToken, siteId, userMail);

      return new OkObjectResult(appRegs);
    }
  }
}
