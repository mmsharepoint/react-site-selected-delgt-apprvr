using Azure.Core;
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Sites.Item;

namespace SampleConsumeFunc.Services
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

    public async Task<bool> UpdateSiteDescreption(string userAssertion, string siteUrl)
    {
      _appGraphClient = GetUserGraphClient(userAssertion);
      Uri uri = new Uri(siteUrl);
      string domain = uri.Host;
      var path = uri.LocalPath;
      var site = await _appGraphClient.Sites[$"{domain}:{path}"].GetAsync();

      var newSite = new Site
      {
        Description = "Next Description"
      };

      await _appGraphClient.Sites[site.Id].PatchAsync(newSite);
      return true;
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
  }
}
