using System;
using System.IO;
using Azure.Identity;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Graph.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using SampleFunction.Services;
using Microsoft.Graph;
using Azure.Core;
using System.Threading;
using Microsoft.Graph.Models.ExternalConnectors;
using System.Collections.Generic;

namespace SampleFunction
{
  public class SampleFunction
  {
    //private readonly ITokenValidationService _tokenValidationService;
    //private readonly IGraphService _graphClientService;
    private readonly ILogger _logger;
    private readonly ILoggerFactory _loggerFactory;

    public SampleFunction(
        //ITokenValidationService tokenValidationService,
        //IGraphService graphClientService,
        ILoggerFactory loggerFactory)
    {
      //_tokenValidationService = tokenValidationService;
      ////_graphClientService = graphClientService;
      _loggerFactory = loggerFactory;
      _logger = loggerFactory.CreateLogger<SampleFunction>();
    }

    [FunctionName("Function1")]
    public async Task<IActionResult> Run(
        [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestData req, 
        ILogger log)
    {
      log.LogInformation("C# HTTP trigger function processed a request.");

      var _config = new ConfigurationBuilder().Build();
      TokenValidationService _tokenValidationService = new TokenValidationService(_config, _loggerFactory);
      var bearerToken = await _tokenValidationService
          .ValidateAuthorizationHeaderAsync(req);

      log.LogInformation("Bootstrap token: " + bearerToken);

      var scopes = new[] { "https://graph.microsoft.com/.default" };

      var tenantId = "7e77d071-ed08-468a-bc75-e8254ba77a21";
      var clientId = "0a8dfbc9-0423-495b-a1e6-1055f0ca69c2";
      var clientSecret = "C9i8Q~kv7UmRZ6Yd9NSE1VSnDZAs6-EF3A6bUa4~";

      // using Azure.Identity;
      var options = new OnBehalfOfCredentialOptions
      {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
      };

      // This is the incoming token to exchange using on-behalf-of flow
      var oboToken = bearerToken;

      OnBehalfOfCredential onBehalfOfCredential = new OnBehalfOfCredential(
          tenantId, clientId, clientSecret, oboToken, options);

      var tokenRequestContext = new TokenRequestContext(scopes);
      string accessToken = onBehalfOfCredential.GetTokenAsync(tokenRequestContext, new CancellationToken()).Result.Token;
      //string accessToken = _graphClientService.GetUserAssessToken(bearerToken);
      log.LogInformation($"Access token: {accessToken}");
    
      // var graphClient = new GraphServiceClient(onBehalfOfCredential, scopes); 

      string responseMessage = $"Hello. This HTTP triggered function executed successfully.";

            


      return new OkObjectResult(responseMessage);
    }
  }
}
