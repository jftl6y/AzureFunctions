using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Microsoft.Identity.Client;
using Newtonsoft.Json;
using System.Net.Http.Headers;

using Microsoft.Graph;

namespace Microsoft.Azure
{
    public static class CreateAADGuestUser
    {
        [FunctionName("CreateAADGuestUser")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            try
            {
                string invitedName = req.Query["name"];
                string invitedEmail = req.Query["email"];
                string graphEndpoint = Environment.GetEnvironmentVariable("graphEndpoint");
                string graphScope = $"{graphEndpoint}/.default";
                string graphUri = $"{graphEndpoint}/v1.0";
                string clientId = Environment.GetEnvironmentVariable("clientId");
                string clientSecret = Environment.GetEnvironmentVariable("clientSecret");
                string tenantId = Environment.GetEnvironmentVariable("tenantId");
                string redirectUri = Environment.GetEnvironmentVariable("redirectUri");
                string loginUri = Environment.GetEnvironmentVariable("loginUri");
                loginUri = $"{loginUri}/{tenantId}/oauth2/v2.0/token";
                var scopes = new[] { graphScope };

                var clientApplication = ConfidentialClientApplicationBuilder
                .Create(clientId)
                .WithTenantId(tenantId)
                .WithClientSecret(clientSecret)
                .WithAuthority(loginUri)
                .Build();

                var authenticationResult = await clientApplication.AcquireTokenForClient(scopes).ExecuteAsync();

                // Create GraphClient and attach auth header to all request (acquired on previous step)
                var graphClient = new GraphServiceClient(graphUri,
                    new DelegateAuthenticationProvider(requestMessage =>
                    {
                        requestMessage.Headers.Authorization =
                            new AuthenticationHeaderValue("bearer", authenticationResult.AccessToken);

                        return Task.FromResult(0);
                    }));


                var invitation = new Invitation
                {
                    InvitedUserDisplayName = invitedName,
                    InvitedUserEmailAddress = invitedEmail,
                    InviteRedirectUrl = redirectUri
                };
                await graphClient.Invitations.Request().AddAsync(invitation);
                var responseMessage = "Success";

                return new OkObjectResult(responseMessage);
            }
            catch (Exception ex)
            {
                return new BadRequestObjectResult(ex.Message);
            }
        }
    }
}
