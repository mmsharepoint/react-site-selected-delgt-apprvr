using Microsoft.Graph;

namespace ApplyPermissions.Services
{
  public interface IGraphService
  {
    public GraphServiceClient? GetUserGraphClient(string userAssertion);
    public GraphServiceClient? GetAppGraphClient();
    public string? GetUserAssessToken(string userAssertion);
  }
}
