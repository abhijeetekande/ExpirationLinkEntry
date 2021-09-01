using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.SharePoint.Client;
using PnP.Framework;
using System.Security;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Auth;
using Microsoft.Azure.Cosmos.Table;
//using Microsoft.WindowsAzure.Storage.Table;

namespace swxExpirationLink
{
    public static class ExpirationLinkEntry
    {
        [FunctionName("ExpirationLinkEntry")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequest req,
            [Table("tblExpirationLinks", Connection = "AzureWebJobsStorage")] CloudTable cloudTable,
            ILogger log)
        {
            log.LogInformation("C# HTTP trigger function processed a request.");
            //IAsyncCollector<ExpirationLinksTableEntity> ExpirationLinksTableCollector
            // string name = req.Query["name"];

            string requestBody = await new StreamReader(req.Body).ReadToEndAsync();
            dynamic data = JsonConvert.DeserializeObject(requestBody);
            //name = name ?? data?.name;
            try
            {
                //    string[] users = Convert.ToString(data.targetUsers).Contains(';') ? Convert.ToString(data.targetUsers).Split(';') : new string[0];
                //    log.LogInformation($"C# {users.Length}");
                //    if (users.Length > 1)
                //{
                //        foreach(string user in users)
                //        {
                //            TableOperation insertOrMergeOperation = TableOperation.InsertOrMerge(new ExpirationLinksTableEntity()
                //            {
                //                PartitionKey = data.currentUser,
                //                RowKey = Convert.ToString(data.targetUsers).Split('@')[0] + data.ItemURL + data.expirationDate,
                //                ItemURL = data.ItemURL,
                //                SharedByUser = data.currentUser,
                //                SharedWithUser = user,
                //                PermissionLevel = Convert.ToBoolean(data.editingEnabled) ? "Edit" : "Read",
                //                ExpirationDate = data.expirationDate,
                //                WebURL = data.webUrl

                //            });
                //            TableResult result = await cloudTable.ExecuteAsync(insertOrMergeOperation);
                //        }

                //}
                //else {
                log.LogInformation($"{String.Concat(Convert.ToString(data.targetUsers).Split('@')[0], Convert.ToString(data.ItemURL))}");
                TableOperation insertOrMergeOperation = TableOperation.InsertOrMerge(new ExpirationLinksTableEntity()
                    {
                        PartitionKey = Guid.NewGuid().ToString(),// Convert.ToString(data.currentUser),  
                        RowKey = Guid.NewGuid().ToString(), //String.Concat(Convert.ToString(data.targetUsers).Split('@')[0], Convert.ToString(data.ItemURL)),// Guid.NewGuid().ToString(),//$"{Convert.ToString(data.targetUsers).Split('@')[0]}{data.ItemURL}",
                        ItemURL = data.ItemURL,
                        SharedByUser = data.currentUser,
                        SharedWithUser = data.targetUsers,
                        PermissionLevel = Convert.ToBoolean(data.editingEnabled) ? "Edit" : "Read",
                        ExpirationDate = data.expirationDate,
                        WebURL = data.webUrl,
                        Expired = Convert.ToBoolean(false),
                        SiteID = data.SiteID,
                        ListID= data.ListID,
                        ItemID = data.ItemID,
                        PermissionID= data.PermissionID

                    });
                    TableResult result = cloudTable.Execute(insertOrMergeOperation);
              //  }          
            }
            catch(Exception ex)
            {
                log.LogInformation($"Error: {ex.StackTrace}");
            }
            try
            {
                //await ExpirationLinksTableCollector.AddAsync(new ExpirationLinksTableEntity()
                //{
                //    PartitionKey = Guid.NewGuid().ToString(),
                //    RowKey = Guid.NewGuid().ToString(),
                //    ItemURL = data.ItemURL,
                //    SharedByUser = data.currentUser,
                //    SharedWithUser = data.targetUsers,
                //    PermissionLevel = Convert.ToBoolean(data.editingEnabled)?"Edit":"Read",
                //    ExpirationDate = data.expirationDate,
                //    WebURL = data.webUrl

                //});
            }
            catch (Exception x)
            {
                log.LogInformation($"Error: {x}");
            }
            string responseMessage = $"Hello, {data.currentUser}. your shared link entered in azure table.";
            log.LogInformation($"C# Site:");

            return new OkObjectResult(responseMessage);
        }

        public class ExpirationLinksTableEntity : TableEntity
        {
            public string ItemURL { get; set; }
            public string SharedByUser { get; set; }
            public string SharedWithUser { get; set; }
            public string PermissionLevel { get; set; }
            public string ExpirationDate { get; set; }
            public string WebURL { get; set; }
            public Boolean Expired { get; set; }
            public string SiteID { get; set; }
            public string ListID { get; set; }
            public string ItemID { get; set; }
            public string PermissionID { get; set; }


            //public ExpirationLinksTableEntity(string sharedByUser, string sharedWithUser, string permissionLevel, DateTime expirationDate)
            //{
            //    SharedByUser = sharedByUser;
            //    SharedWithUser = sharedWithUser;
            //    PermissionLevel = permissionLevel;
            //    ExpirationDate = expirationDate;

            //}
            //public ExpirationLinksTableEntity() { }
        }
    }
}
