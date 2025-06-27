using IC_Loader_Pro.Models;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using static IC_Loader_Pro.Module1;

namespace IC_Loader_Pro.Services
{
    public class GraphApiService
    {
        private readonly GraphServiceClient _graphClient;

        // Configuration for your app registration in Azure
        private const string ClientId = "YOUR_CLIENT_ID_HERE"; // Must be replaced
        private const string TenantId = "YOUR_TENANT_ID_HERE"; // Can often be "common" or your organization's tenant ID
        private static readonly string[] _scopes = { "User.Read", "Mail.Read" };

        public GraphApiService(GraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }

        /// <summary>
        /// Creates and initializes a GraphServiceClient, handling authentication.
        /// </summary>
        public static async Task<GraphApiService> CreateAsync()
        {
            var publicClientApp = PublicClientApplicationBuilder
                .Create(ClientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, TenantId)
                .WithDefaultRedirectUri()
                .Build();

            AuthenticationResult authResult;
            var accounts = await publicClientApp.GetAccountsAsync();

            try
            {
                // Try to get a token silently
                authResult = await publicClientApp.AcquireTokenSilent(_scopes, accounts.FirstOrDefault()).ExecuteAsync();
            }
            catch (MsalUiRequiredException)
            {
                // If silent fails, show the interactive login
                authResult = await publicClientApp.AcquireTokenInteractive(_scopes).ExecuteAsync();
            }

            var graphClient = new GraphServiceClient(new DelegateAuthenticationProvider((requestMessage) =>
            {
                requestMessage.Headers.Authorization =
                    new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authResult.AccessToken);
                return Task.CompletedTask;
            }));

            return new GraphApiService(graphClient);
        }

        /// <summary>
        /// Gets emails from a folder specified by a full path, applying a three-state test filter.
        /// </summary>
        public async Task<List<EmailItem>> GetEmailsFromFolderPathAsync(string fullFolderPath, string testSenderEmail, bool? isInTestMode)
        {
            if (string.IsNullOrWhiteSpace(fullFolderPath) || !fullFolderPath.StartsWith("\\\\"))
            {
                throw new ArgumentException("Invalid folder path format. Path must start with '\\\\'.", nameof(fullFolderPath));
            }

            var results = new List<EmailItem>();
            try
            {
                var (storeName, folderPath) = ParseOutlookPath(fullFolderPath);

                // Get the folder ID from its path
                string targetFolderId = await GetFolderIdFromPathAsync(folderPath);
                if (string.IsNullOrEmpty(targetFolderId))
                {
                    throw new System.IO.DirectoryNotFoundException($"The Outlook folder specified by the path '{fullFolderPath}' could not be found.");
                }

                // Fetch messages from the folder
                var messages = await _graphClient.Me.MailFolders[targetFolderId].Messages
                                     .Request()
                                     .Select("subject,receivedDateTime,sender,from,hasAttachments,internetMessageId")
                                     .GetAsync();

                foreach (var message in messages)
                {
                    results.Add(new EmailItem
                    {
                        Emailid = message.InternetMessageId,
                        Subject = message.Subject,
                        ReceivedTime = message.ReceivedDateTime?.DateTime ?? DateTime.MinValue,
                        SenderName = message.Sender?.EmailAddress?.Name,
                        SenderEmailAddress = message.Sender?.EmailAddress?.Address,
                        AttachmentCount = message.HasAttachments == true ? 1 : 0 // Note: Graph does not give a direct count here
                    });
                }
            }
            catch (Exception ex)
            {
                Log.recordError($"Failed to retrieve emails using Graph API for path '{fullFolderPath}'.", ex, "GetEmailsFromFolderPathAsync");
                throw;
            }

            // Apply the same three-state filtering logic
            if (!isInTestMode.HasValue || string.IsNullOrWhiteSpace(testSenderEmail))
            {
                return results;
            }

            if (isInTestMode.Value)
            {
                return results.Where(e => e.SenderEmailAddress.Equals(testSenderEmail, StringComparison.OrdinalIgnoreCase)).ToList();
            }
            else
            {
                return results.Where(e => !e.SenderEmailAddress.Equals(testSenderEmail, StringComparison.OrdinalIgnoreCase)).ToList();
            }
        }

        /// <summary>
        /// Finds the ID of a mail folder given its path (e.g., "Inbox/Subfolder").
        /// </summary>
        private async Task<string> GetFolderIdFromPathAsync(string path)
        {
            var folderNames = path.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);

            // Start with the root folder collection
            var currentFolders = await _graphClient.Me.MailFolders.Request().GetAsync();
            string currentFolderId = null;

            foreach (var folderName in folderNames)
            {
                var foundFolder = currentFolders.FirstOrDefault(f => f.DisplayName.Equals(folderName, StringComparison.OrdinalIgnoreCase));
                if (foundFolder == null) return null; // Path not found

                currentFolderId = foundFolder.Id;
                // Get the children of the current folder for the next iteration
                currentFolders = await _graphClient.Me.MailFolders[currentFolderId].ChildFolders.Request().GetAsync();
            }

            return currentFolderId;
        }

        private (string storeName, string folderPath) ParseOutlookPath(string fullPath)
        {
            var parts = fullPath.TrimStart('\\').Split(new[] { '\\' }, 2);
            if (parts.Length < 2)
            {
                throw new ArgumentException("Path must include at least a store name and a folder name.", nameof(fullPath));
            }
            // Note: With Graph, the store name is implied (it's always the user's primary mailbox)
            // but we parse it to maintain compatibility with the existing path format.
            return (parts[0], parts[1]);
        }
    }
}