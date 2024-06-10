using Microsoft.Azure.Functions.Worker.Http;

namespace ApplyPermissions.Services
{
    public interface ITokenValidationService
    {
        public Task<string?> ValidateAuthorizationHeaderAsync(
            HttpRequestData request);
    }
}
