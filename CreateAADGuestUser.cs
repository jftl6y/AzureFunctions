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
                                SendInvitationMessage = true
                            };
                            users.Add(user);
                        }
                    }
                    loopCounter++;
                }

                List<IActionResult> results = new List<IActionResult>();
                foreach (var userInfo in users)
                {
                    results.Add(await GraphClientMethods.SendUserInfo(userInfo));
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
        
    }

}
