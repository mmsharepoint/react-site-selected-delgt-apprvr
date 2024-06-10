using Microsoft.Graph;

namespace SampleFunction.Services
{
  public interface IGraphService
  {
    public GraphServiceClient? GetUserGraphClient(string userAssertion);
    public GraphServiceClient? GetAppGraphClient();
    public string? GetUserAssessToken(string userAssertion);
  }
}
