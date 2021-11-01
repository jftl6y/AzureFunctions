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
using System.Collections.Generic;
using System.Linq;

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
                string inputCsv = await new StreamReader(req.Body).ReadToEndAsync();

                var inputRows = inputCsv.Split("\r\n");
                var users = new List<UserInfo>();
                int loopCounter = 0;

                foreach (var row in inputRows)
                {
                    if (loopCounter != 0) //First row has headers
                    {
                        List<string> rowColumns = new List<string>(row.Split(","));
                        if (rowColumns.Count == 5) //Check and see that we have the expected number of columns
                        {
                            var user = new UserInfo()
                            {
                                LastName = rowColumns[0],
                                FirstName = rowColumns[1],
                                CompanyName = rowColumns[3],
                                Email = rowColumns[4],
                                SendInvitationMessage = false
                            };
                            users.Add(user);
                        }
                    }
                    loopCounter++;
                }

                List<IActionResult> results = new List<IActionResult>();
                foreach (var userInfo in users)
                {
                    results.Add(await SendUserInfo(userInfo));
                }
                var failedObjectCount = results.Where(r => r.GetType().ToString() == "Microsoft.AspNetCore.Mvc.BadRequestObjectResult").Count();
                var successObjectCount = results.Where(r => r.GetType().ToString() == "Microsoft.AspNetCore.Mvc.OkObjectResult").Count();

                return new OkObjectResult($"Processed {successObjectCount} records successfully, {failedObjectCount} records failed");
            }
            catch (Exception ex)
            {
                return new BadRequestObjectResult(ex.Message);
            }
        }
        public static async Task<IActionResult> SendUserInfo(UserInfo userInfo)
        {
            try
            {
                string graphEndpoint = Environment.GetEnvironmentVariable("graphEndpoint");
                string graphScope = $"{graphEndpoint}/.default";
                string graphUri = $"{graphEndpoint}/v1.0";
                string clientId = Environment.GetEnvironmentVariable("clientId");
                string clientSecret = Environment.GetEnvironmentVariable("clientSecret");
                string tenantId = Environment.GetEnvironmentVariable("tenantId");
                string redirectUri = Environment.GetEnvironmentVariable("redirectUri");
                string loginUri = Environment.GetEnvironmentVariable("loginUri");
                string tenantDomain = Environment.GetEnvironmentVariable("tenantDomain");
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
                var userDisplayName = String.IsNullOrEmpty(userInfo.DisplayName) ? $"{userInfo.FirstName} {userInfo.LastName}" : userInfo.DisplayName;

                var invitation = new Invitation
                {
                    InvitedUserDisplayName = userDisplayName,
                    InvitedUserEmailAddress = userInfo.Email,
                    InviteRedirectUrl = redirectUri,
                    SendInvitationMessage = userInfo.SendInvitationMessage
                };
                await graphClient.Invitations.Request().AddAsync(invitation);
                
                //var users = graphClient.Users.Request().GetAsync().Result;


                var emailAddressParts = userInfo.Email.Split("@");
                var upn = $"{emailAddressParts[0]}_{emailAddressParts[1]}#EXT#@{tenantDomain}";
                
                var updateUser = new User
                {
                    CompanyName = userInfo.CompanyName
                };
                
                await graphClient.Users[upn].Request().UpdateAsync(updateUser); //todo fix why this isn't finding the user
                
                var responseMessage = "Success";

                return new OkObjectResult(responseMessage);
            }
            catch (Exception ex)
            {
                return new BadRequestObjectResult(ex.Message);
            }
        }
        public class UserInfo
        {
            public string LastName { get; set; }
            public string FirstName { get; set; }
            public string DisplayName { get; set; }
            public string CompanyName { get; set; }
            public string Email { get; set; }
            public bool SendInvitationMessage { get; set; }
        }
    }

}
