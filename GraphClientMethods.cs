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
    public class GraphClientMethods {
        public static async Task<IActionResult> SendUserInfo(UserInfo userInfo)
        {
            try
            {
                string redirectUri = Environment.GetEnvironmentVariable("redirectUri");

                var graphClient = await GetGraphServiceClient();
                var userDisplayName = String.IsNullOrEmpty(userInfo.DisplayName) ? $"{userInfo.FirstName} {userInfo.LastName}" : userInfo.DisplayName;

                var invitation = new Invitation
                {
                    InvitedUserDisplayName = userDisplayName,
                    InvitedUserEmailAddress = userInfo.Email,
                    InviteRedirectUrl = redirectUri,
                    SendInvitationMessage = userInfo.SendInvitationMessage
                };
                var ret = await graphClient.Invitations.Request().AddAsync(invitation);

                string userId = null;
                int loopCounter = 0;
                while (String.IsNullOrEmpty(userId) && loopCounter < 10000) //todo figure out a better way to anticipate the creation of the user object
                {
                    var filterString = "";  //todo fix this filter string to pre-filter results
                    var users = graphClient.Users.Request().Filter(filterString).GetAsync().Result;
                    var filteredUser = users.Where(u => u.Mail == userInfo.Email).FirstOrDefault();
                    if (filteredUser != null) { userId = filteredUser.Id; }
                    loopCounter++;
                }
                var user = new User
                {
                    CompanyName = userInfo.CompanyName
                };
                await graphClient.Users[userId].Request().UpdateAsync(user); 
                var responseMessage = "Success";

                return new OkObjectResult(responseMessage);
            }
            catch (Exception ex)
            {
                return new BadRequestObjectResult(ex.Message);
            }
        }
        private static async Task<GraphServiceClient> GetGraphServiceClient()
        {

            string graphEndpoint = Environment.GetEnvironmentVariable("graphEndpoint");
            string graphScope = $"{graphEndpoint}/.default";
            string graphUri = $"{graphEndpoint}/v1.0";
            string clientId = Environment.GetEnvironmentVariable("clientId");
            string clientSecret = Environment.GetEnvironmentVariable("clientSecret");
            clientId = "663991fb-a753-4a9b-b97d-0b508648d589";
            clientSecret = "RV07Q~6k2K205DL9tNGfed2t_hmgFO6DmyI7R";

            string tenantId = Environment.GetEnvironmentVariable("tenantId");
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
            return new GraphServiceClient(graphUri,
                 new DelegateAuthenticationProvider(requestMessage =>
                 {
                     requestMessage.Headers.Authorization =
                         new AuthenticationHeaderValue("bearer", authenticationResult.AccessToken);

                     return Task.FromResult(0);
                 }));
        }
    }
}