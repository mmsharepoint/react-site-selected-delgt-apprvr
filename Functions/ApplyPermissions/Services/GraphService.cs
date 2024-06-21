using ApplyPermissions.Model;
using Azure;
using Azure.Core;
using Azure.Identity;
using Google.Protobuf.WellKnownTypes;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Identity.Client;
using static Microsoft.Graph.CoreConstants;

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

    public async Task<bool> ApplySitePermission(string userAssertion, string siteUrl, string role, string appId)
    {
      _appGraphClient = GetUserGraphClient(userAssertion);
      Uri uri = new Uri(siteUrl);
      string domain = uri.Host;
      var path = uri.LocalPath;
      var site = await _appGraphClient.Sites[$"{domain}:{path}"].GetAsync();
      string appDisplayName = await this.getServicePrincipal(appId);

      var requestBody = new Permission
      {
        Roles = new List<string>
        {
          role,
        },
        GrantedToIdentities = new List<IdentitySet>
        {
          new IdentitySet
          {
            Application = new Identity
            {
              Id = appId,
              DisplayName = appDisplayName
            },
          },
        },
      };

      try
      {
        await _appGraphClient.Sites[site.Id].Permissions.PostAsync(requestBody);
      }
      catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
      {
        _logger.LogError(ex.Message);
      }

      return true;
    }

    public async Task<List<AppReg>> GetServicePrincipals(string userAssertion, string prefix)
    {
      _appGraphClient = GetUserGraphClient(userAssertion);
      string filter = $"startswith(DisplayName,'{prefix}')";
      var servicePrincipals = await _appGraphClient.ServicePrincipals.GetAsync(requestConfig =>
      {
        requestConfig.QueryParameters.Filter = filter;
      });
      List<AppReg> appRegs = new List<AppReg>();
      servicePrincipals.Value.ForEach(service =>
      {
        appRegs.Add(new AppReg { Id = service.Id, DisplayName = service.DisplayName });
      });
      return appRegs;
    }

    public async Task<List<SearchSite>> SearchSites(string userAssertion, string queryText)
    {
      _appGraphClient = GetUserGraphClient(userAssertion);
      var sites = await _appGraphClient.Sites.GetAsync((requestConfiguration) =>
      {
        requestConfiguration.QueryParameters.Search = queryText;
      });
      List<SearchSite> siteResults = new List<SearchSite>();
      sites.Value.ForEach(site =>
      {
        siteResults.Add(new SearchSite { Id = site.Id, Title = site.DisplayName, Url = site.WebUrl });
      });
      return siteResults;
    }

    public async Task<bool> IsSiteAdmin(string userAssertion, string siteId, string userMail)
    {
      _appGraphClient = GetUserGraphClient(userAssertion);
      string filter = $"fields/EMail eq '{userMail}'";
      var user = await _appGraphClient.Sites[siteId].Lists["User Information List"].Items.GetAsync(requestConfig =>
      {
        requestConfig.Headers.Add("Prefer", @"HonorNonIndexedQueriesWarningMayFailRandomly");
        requestConfig.QueryParameters.Expand = ["fields"];
        requestConfig.QueryParameters.Select = ["fields"];
        requestConfig.QueryParameters.Filter = filter;
      });
      bool val = user.Value[0].Fields.AdditionalData.TryGetValue("IsSiteAdmin", out object isAdmin);
      if (val)
      {
        return bool.Parse(isAdmin.ToString());
      }
      else
      {
        return false;
      }
    }

    public GraphServiceClient? GetUserGraphClient(string userAssertion)
    {
      var tenantId = _config["tenantId"];
      var clientId = _config["clientId"];
      var clientSecret = _config["clientSecret"];
      var scopes = new[] { "https://graph.microsoft.com/.default" };

      if (string.IsNullOrEmpty(tenantId) ||
          string.IsNullOrEmpty(clientId) ||
          string.IsNullOrEmpty(clientSecret))
      {
        _logger.LogError("Required settings missing: 'tenantId', 'clientId', or 'clientSecret'.");
        return null;
      }

      var onBehalfOfCredential = new OnBehalfOfCredential(
          tenantId, clientId, clientSecret, userAssertion);

      return new GraphServiceClient(onBehalfOfCredential, scopes);
    }

    private async Task<string> getServicePrincipal (string appId)
    {
      string displayName = String.Empty;
      try
      {
        var servicePrincipal = await _appGraphClient.ServicePrincipals[appId].GetAsync();
        displayName = servicePrincipal.DisplayName;
      }
      catch (Microsoft.Graph.Models.ODataErrors.ODataError ex)
      {
        _logger.LogError(ex.Message);
      }

      return displayName;     
  }
}
}
