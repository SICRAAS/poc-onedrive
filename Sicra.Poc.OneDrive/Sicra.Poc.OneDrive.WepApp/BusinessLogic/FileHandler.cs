using System;
using System.IO;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;

namespace Sicra.Poc.OneDrive.WepApp.BusinessLogic
{
    public static class FileHandler
    {
        
        private const string TokenScheme = "Bearer";
        private static readonly string[] Scopes = new[] { "https://graph.microsoft.com/.default" };
        
        
        
        
        /// <summary>
        /// Store file in folder on Sharepoint Site with Public Read Access
        /// </summary>
        /// <param name="fileBytes"></param>
        /// <param name="siteId"></param>
        /// <param name="folderName"></param>
        /// <returns>Public Url to the file</returns>
        public static async Task<string> StoreOnSharepointSite(byte[] fileBytes, string siteId, string folderName)
        {
            // DefaultAzureCredential reads the following environment variables
            // AZURE_TENANT_ID      - The AzureAd/EntraId Tenant Id
            // AZURE_CLIENT_ID      - Service Principal (Enterprise application) CLient Id
            // AZURE_CLIENT_SECRET  - Service Principal Client Secret
            
           
            // Get Access Token
            var credentials = new DefaultAzureCredential(new DefaultAzureCredentialOptions());
            var token = await credentials.GetTokenAsync(new TokenRequestContext(Scopes));
            // The token is valid for 3900 seconds (65 minutes) and should be stored in a MemoryCache.
            
            
            // Create a graph client and add the access token to the client
            var graphClient = new GraphServiceClient(
                new DelegateAuthenticationProvider(
                    (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue(TokenScheme, token.Token);
                        return Task.FromResult(0);
                    }));

            
            // Create a random filename
            var destinationFilename = $"{Guid.NewGuid()}.xlsx";
            
            
            // Store file
            var memoryStream = new MemoryStream(fileBytes);
            var file = await graphClient
                .Sites[siteId]
                .Drive
                .Root.ItemWithPath($"{folderName}/{destinationFilename}")
                .Content
                .Request()
                .PutAsync<DriveItem>(memoryStream);

            
            // Create link with anonymous view permission that is valid for ten days
            var permissionExpireTimestamp = DateTimeOffset.UtcNow.AddDays(10);
            var permission = await graphClient
                .Sites[siteId]
                .Drive
                .Items[file.Id]
                .CreateLink("view","anonymous",permissionExpireTimestamp,null,null)
                .Request()
                .PostAsync();

            
            // Return url
            return permission.Link.WebUrl;
        }
        
        
        
        
        /// <summary>
        /// Reads file from App_Data folder
        /// </summary>
        /// <param name="filename"></param>
        /// <returns>ByteArray of file content</returns>
        public static byte[] ReadMockFile(string filename)
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var fileBytes = System.IO.File.ReadAllBytes($"{baseDir}/App_Data/{filename}");
            return fileBytes;
        }
    }
}