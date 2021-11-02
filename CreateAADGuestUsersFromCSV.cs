
using System.IO;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;
using Microsoft.AspNetCore.Mvc;
using System.Collections.Generic;
using System.Linq;
using Azure.Storage.Blobs;

using Microsoft.Graph;

namespace Microsoft.Azure
{
    public static class CreateAADGuestUsersFromCSV
    {
        [FunctionName("CreateAADGuestUsersFromCSV")]
        public static void Run([BlobTrigger("users/{name}", Connection = "UserCsvStorage")]Stream myBlob, string name, ILogger log)
        {
            log.LogInformation($"C# Blob trigger function Processed blob\n Name:{name} \n Size: {myBlob.Length} Bytes");
            
            myBlob.Position = 0;
            
            string inputCsv = new StreamReader(myBlob).ReadToEnd();

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
                    results.Add(GraphClientMethods.SendUserInfo(userInfo).Result);
                }
                var failedObjectCount = results.Where(r => r.GetType().ToString() == "Microsoft.AspNetCore.Mvc.BadRequestObjectResult").Count();
                var successObjectCount = results.Where(r => r.GetType().ToString() == "Microsoft.AspNetCore.Mvc.OkObjectResult").Count();

                log.LogInformation($"Processed {successObjectCount} records successfully, {failedObjectCount} records failed");

                //todo Define and implement post-processing blob cleanup
                

        }
    }

    
}
