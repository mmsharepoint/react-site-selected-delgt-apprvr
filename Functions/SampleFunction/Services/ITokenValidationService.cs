using Microsoft.Azure.Functions.Worker.Http;

namespace SampleFunction.Services
{
    public interface ITokenValidationService
    {
        public Task<string?> ValidateAuthorizationHeaderAsync(
            HttpRequestData request);
    }
}
