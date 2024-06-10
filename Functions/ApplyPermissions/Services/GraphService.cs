using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;

namespace ApplyPermissions.Services
{
  public class GraphService : IGraphService
  {
    private readonly IConfiguration _config;
    private readonly ILogger _logger;
    private GraphServiceClient? _appGraphClient;

    public GraphService(IConfiguration config, ILogger<GraphService> logger)
    {
      _config = config;
      _logger = logger;
    }

    public GraphServiceClient? GetUserGraphClient(string userAssertion)
    {
      var tenantId = _config["tenantId"];
      var clientId = _config["clientId"];
      var clientSecret = _config["clientSecret"];

      if (string.IsNullOrEmpty(tenantId) ||
          string.IsNullOrEmpty(clientId) ||
          string.IsNullOrEmpty(clientSecret))
      {
        _logger.LogError("Required settings missing: 'tenantId', 'clientId', and 'clientSecret'.");
        return null;
      }

      var onBehalfOfCredential = new OnBehalfOfCredential(
          tenantId, clientId, clientSecret, userAssertion);

      return new GraphServiceClient(onBehalfOfCredential);
    }

    public string GetUserAssessToken(string userAssertion)
    {
      var scopes = new[] { "https://graph.microsoft.com/.default" };
      var tenantId = _config["tenantId"];
      var clientId = _config["clientId"];
      var clientSecret = _config["clientSecret"];

      if (string.IsNullOrEmpty(tenantId) ||
          string.IsNullOrEmpty(clientId) ||
          string.IsNullOrEmpty(clientSecret))
      {
        _logger.LogError("Required settings missing: 'tenantId', 'clientId', and 'clientSecret'.");
        return null;
      }

      var onBehalfOfCredential = new OnBehalfOfCredential(
          tenantId, clientId, clientSecret, userAssertion);
      var tokenRequestContext = new TokenRequestContext(scopes);
      string accessToken = onBehalfOfCredential.GetTokenAsync(tokenRequestContext, new CancellationToken()).Result.Token;
      return accessToken;
    }

    public GraphServiceClient GetAppGraphClient()
    {
      var scopes = new[] { "https://graph.microsoft.com/.default" };

      var tenantId = _config["tenantId"];
      var clientId = _config["clientId"];
      var clientSecret = _config["apiClientSecret"];

      if (string.IsNullOrEmpty(tenantId) ||
          string.IsNullOrEmpty(clientId) ||
          string.IsNullOrEmpty(clientSecret))
      {
        _logger.LogError("Required settings missing: 'tenantId', 'apiClientId', and 'apiClientSecret'.");
        return null;
      }
      var options = new ClientCertificateCredentialOptions
      {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
      };

      var clientCertCredential = new ClientCertificateCredential(
          tenantId, clientId, clientSecret, options);

      var graphClient = new GraphServiceClient(clientCertCredential, scopes);
      return graphClient;
    }
  }
}
