using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.Extensions.Logging;
using Microsoft.SharePoint.Client;
using PnP.Framework;
using System.Security;
using Microsoft.Identity.Client;
using Microsoft.Azure.ActiveDirectory.GraphClient;
using Microsoft.Azure.Services.AppAuthentication;
using User = Microsoft.SharePoint.Client.User;
using Microsoft.Azure.Cosmos.Table;
using Microsoft.IdentityModel.Clients.ActiveDirectory;
using System.Security.Cryptography.X509Certificates;
using Microsoft.Extensions.Logging;
using System.Threading.Tasks;
using System.IO;
using PnP.Core.Auth;

namespace swxExpirationLink
{
    public static class PermissionRemover
    {
        [FunctionName("PermissionRemover")]
        public static void Run([TimerTrigger("0 */10 * * * *")]TimerInfo myTimer, [Table("tblExpirationLinks", Connection = "AzureWebJobsStorage")] CloudTable cloudTable, ILogger log)
        {
           
            try
            {
                log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
                // const string Storageaccountname = "storageaccountswxab8579";
                // const string StorageAccountKey = "vdgwCvlo0iJho61MMb8jLlIe2MeIUtnAWHie9MvdZlj6qdxcrnN2CRsJ8XzY7wdaq2hoUcV8dNNVjkjIgWc/2g==";
                //  var storageAccount = new CloudStorageAccount(new StorageCredentials(Storageaccountname, StorageAccountKey), false);
                DateTime accMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                string formatted = accMonth.ToLongTimeString();
               
                TableQuery<ExpirationLinksTableEntity> query = new TableQuery<ExpirationLinksTableEntity>()
                    .Where(TableQuery.GenerateFilterCondition("ExpirationDate", QueryComparisons.Equal, "Wed Aug 11 2021 00:00:00 GMT+0530 (India Standard Time)" )); //DateTime.Now.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss.fffffffK")

                //var segment = cloudTable.ExecuteQuery(query,null);// ExecuteQuerySegmentedAsync(query, null);
                foreach (ExpirationLinksTableEntity entity in cloudTable.ExecuteQuery(query))
                {
                    Console.WriteLine($"Day: {entity.PartitionKey}, ID:{entity.RowKey}\tName:{entity.SharedByUser}\tDescription{entity.SharedWithUser}\tWebURL:{entity.WebURL}");
                    RemoveSPUserPermission(entity.SharedWithUser, entity.ItemURL, entity.WebURL, log);
                }
                //var data = segment.Select(TodoExtensions.ToTodo);
                // var expirationTable = storageAccount.CreateCloudTableClient().GetTableReference("tblExpirationLinks");
                log.LogInformation($"Connected to the table:");
               // var result = expirationTable.ExecuteAsync(TableOperation.Retrieve<ExpirationLinksTableEntity>("abhi", "swayam"));
            }
            catch(Exception x)
            {
                log.LogInformation($"Error: {x}");
            }
           
           // RemoveSPUserPermission(log);
        }

        public class ExpirationLinksTableEntity : TableEntity
        {
            public string ItemURL { get; set; }
            public string SharedByUser { get; set; }
            public string SharedWithUser { get; set; }
            public string PermissionLevel { get; set; }
            public DateTime ExpirationDate { get; set; }
            public string WebURL { get; set; }

            //public ExpirationLinksTableEntity(string sharedByUser, string sharedWithUser, string permissionLevel, DateTime expirationDate)
            //{
            //    SharedByUser = sharedByUser;
            //    SharedWithUser = sharedWithUser;
            //    PermissionLevel = permissionLevel;
            //    ExpirationDate = expirationDate;

            //}
            //public ExpirationLinksTableEntity() { }
        }

      
        //public static void RemoveSPUserPermission(string sharedWithUser, string itemURL,ILogger log)
            public static void RemoveSPUserPermission(string sharedWithUser, string itemUrl, string webURL,ILogger log)
        {
            //var url = Environment.GetEnvironmentVariable("tenantRootUrl");
            //var thumbprint = Environment.GetEnvironmentVariable("certificateThumbprint");
            //var resourceUri = Environment.GetEnvironmentVariable("resourceUri");
            //var authorityUri = Environment.GetEnvironmentVariable("authorityUri");
            //var clientId = Environment.GetEnvironmentVariable("clientId");
            //var ac = new AuthenticationContext(authorityUri, false);
            //var cert = GetCertificate(thumbprint);  //this is the utility method called out above
            //Microsoft.IdentityModel.Clients.ActiveDirectory.ClientAssertionCertificate cac = new Microsoft.IdentityModel.Clients.ActiveDirectory.ClientAssertionCertificate(clientId, cert);
            //var authResult = ac.AcquireTokenAsync(resourceUri, cac).Result;

            //# next section makes calls to SharePoint Online but could easily be to another resource
            //using (ClientContext cc = new ClientContext(url))
            //{
            //    cc.ExecutingWebRequest += (s, e) =>
            //    {
            //        e.WebRequestExecutor.RequestHeaders["Authorization"] = "Bearer " + authResult.AccessToken;
            //    };

            //# make calls through the client context object

            // }
            ///////////////////////////////////////////////////////////////////
            ///
            log.LogInformation("In permission remover");
            var thumbprint = Environment.GetEnvironmentVariable("certificateThumbprint");
            string aadApplicationId = "dd712a30-caac-45eb-b048-1cf1853b8dd8";
            string tenantName = "spadeworxsoftwareservices";
            string sharePointUrl = webURL;// $"https://{tenantName}.sharepoint.com/";
            Web currentWeb;
            log.LogInformation($"Thumbprint : {thumbprint}");
            AuthenticationManager auth = new AuthenticationManager(aadApplicationId, GetCertificate(thumbprint), $"{tenantName}.onmicrosoft.com");

            using (ClientContext ctx = auth.GetContext(sharePointUrl))
            {
                currentWeb = ctx.Web;
                ctx.Load(currentWeb);
                ctx.ExecuteQueryRetry();
          //  }

            log.LogInformation($"Web's title : {currentWeb.Title}");

       /////////////////////////////////////////////////////
       // string SiteUrl = "https://abhiekande.sharepoint.com/sites/dev3"; // "https://spadeworxsoftwareservices.sharepoint.com/teams/abhijit";
            //var clientId = "dd712a30-caac-45eb-b048-1cf1853b8dd8";
           // var clientSecret = "-lpL_5AwAYf.K01VIZS-_s7q1zT95RBeCQ";
           // Web currentWeb;
           // var pwd = "Passw@rd1010$";
           // var username = "abhiekande@abhiekande.onmicrosoft.com";

           // var securePassword = new SecureString();
           // foreach (char c in pwd)
           // {
           //     securePassword.AppendChar(c);
           // }
            
           // AuthenticationManager auth = new AuthenticationManager(username, securePassword);
            //var azureServiceTokenProvider = new AzureServiceTokenProvider();

            // ClientContext context = authManager.GetContextAsync(SiteUrl);// .GetSharePointOnlineAuthenticatedContextTenant(SiteUrl, username, pwd);
           // using (ClientContext ctx = auth.GetContext(SiteUrl))
           // {
                currentWeb = ctx.Web;
                ctx.Load(currentWeb);
                //ctx.ExecuteQueryRetry();
                ctx.Load(ctx.Web, a => a.Lists);
                ctx.ExecuteQueryRetry();

                List list = ctx.Web.Lists.GetByTitle("Documents");
                ctx.ExecuteQueryRetry();
                log.LogInformation("Lib found");
                string document = "log.txt";
                CamlQuery camlQuery = new CamlQuery
                {
                    ViewXml = @"<View><Query><Where><Eq><FieldRef Name='FileLeafRef' /><Value Type='File'>" + document + @"</Value></Eq></Where> 
               </Query> 
                <ViewFields><FieldRef Name='FileRef' /><FieldRef Name='FileLeafRef' /></ViewFields> 
         </View>"
                };
                ctx.Load(list);
                log.LogInformation("List loaded");
                var items = list.GetItems(camlQuery);
                ctx.Load(items);
                ctx.ExecuteQuery();
                log.LogInformation("Items loaded");
                ctx.ExecuteQueryRetry();
                log.LogInformation($"Items: {items.Count}");
                foreach (var item in items)
                {
                    //var user_group = web.SiteGroups.GetByName("Site Members");
                    User user = currentWeb.EnsureUser("swayam@abhiekande.onmicrosoft.com");
                    ctx.Load(user);
                    ctx.ExecuteQueryRetry();
                    var user_group = currentWeb.SiteUsers.GetByLoginName("i:0#.f|membership|"+ sharedWithUser);
                    ctx.Load(user_group);
                    log.LogInformation($"{sharedWithUser} ");
                    ctx.Load(item.RoleAssignments);
                    ctx.ExecuteQueryRetry();

                    foreach (var assignments in item.RoleAssignments)
                    {
                        ctx.Load(assignments.Member);
                        ctx.ExecuteQueryRetry();
                        if (assignments.Member.LoginName == user_group.LoginName)
                        {
                            item.RoleAssignments.GetByPrincipal(user_group).DeleteObject();
                            ctx.ExecuteQueryRetry();
                        }
                    }

                }
            }
            string responseMessage = $"Hello from {currentWeb.Title}";
            log.LogInformation(responseMessage);
        }

        public static X509Certificate2 GetCertificate(string thumbprint)
        {
            var auth = new X509CertificateAuthenticationProvider("dd712a30-caac-45eb-b048-1cf1853b8dd8", "973b3f01-ab0a-4620-b906-eb32095e50cc", StoreName.My, StoreLocation.CurrentUser, thumbprint);

            var bytes = System.IO.File.ReadAllBytes("C:\\swxExpirationLink\\slnSwxExpirationLink\\swxExpirationLink\\Tools\\ShareLinkProj.pfx");
            var cert = new X509Certificate2(bytes);
           // X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            try
            {
                // store.Open(OpenFlags.ReadOnly);

                //   var col = store.Certificates.Find(X509FindType.FindByThumbprint, "908D0E3AE27F7D82C20BD2079D43596CD94A8711", false);
                //  if (col == null || col.Count == 0)
                //   {
                //      return null;
                //  }
                // return col[0];
                return cert;
            }
            finally
            {
               // store.Close();
            }
        }
    }

}
