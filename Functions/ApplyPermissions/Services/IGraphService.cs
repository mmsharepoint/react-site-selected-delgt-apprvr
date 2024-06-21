using ApplyPermissions.Model;
using Microsoft.Graph;

namespace ApplyPermissions.Services
{
  public interface IGraphService
  {
    public GraphServiceClient? GetUserGraphClient(string userAssertion);
    public Task<bool> ApplySitePermission(string userAssertion, string siteUrl, string role, string appId);
    public Task<List<AppReg>> GetServicePrincipals(string userAssertion, string prefix);
    public Task<List<SearchSite>> SearchSites(string userAssertion, string queryText);
    public Task<bool> IsSiteAdmin(string userAssertion, string siteId, string userMail);
  }
}
