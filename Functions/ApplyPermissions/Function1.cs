using Azure.Core;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using ApplyPermissions.Services;
using System.Text;

namespace ApplyPermissions
{
  public class Function1
  {
    private readonly ITokenValidationService _tokenValidationService;
    private readonly IGraphService _graphClientService;
    private readonly ILogger<Function1> _logger;
    private readonly IConfiguration _config;

    public Function1(ITokenValidationService tokenValidationService, IGraphService graphClientService, ILogger<Function1> logger, IConfiguration config)
    {
      _tokenValidationService = tokenValidationService;
      _graphClientService = graphClientService;
      _logger = logger;
      _config = config;
    }

    [Function("Function1")]
    public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req)
    {
      _logger.LogInformation("C# HTTP trigger function processed a request.");
      var bearerToken = await _tokenValidationService
          .ValidateAuthorizationHeaderAsync(req);

      _logger.LogInformation("Bootstrap token: " + bearerToken);

      // This is the incoming token to exchange using on-behalf-of flow
      var oboToken = bearerToken;

      string accessToken = _graphClientService.GetUserAssessToken(bearerToken);
      _logger.LogInformation($"Access token: {accessToken}");
      string documentContents;
      using (Stream receiveStream = req.Body)
      {
        using (StreamReader readStream = new StreamReader(receiveStream, Encoding.UTF8))
        {
          documentContents = readStream.ReadToEndAsync().Result;
        }
      }
      _logger.LogInformation($"Body: {documentContents}");
      return new OkObjectResult($"Welcome to Azure Functions, {req.Query["name"]}!");
    }
  }
}
