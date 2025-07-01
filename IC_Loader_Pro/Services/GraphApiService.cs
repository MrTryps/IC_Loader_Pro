using IC_Loader_Pro.Models;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Authentication;
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

        private GraphApiService(GraphServiceClient graphClient)
        {
            _graphClient = graphClient;
        }

        /// <summary>
        /// Creates and initializes a GraphServiceClient, handling authentication.
        /// </summary>
        public static async Task<GraphApiService> CreateAsync()
        {
            var authProvider = new PublicClientAuthenticationProvider(ClientId, TenantId, _scopes);

            // You will need to add a using statement for System.Net.Http
            var httpClient = new HttpClient();

            var graphClient = new GraphServiceClient(httpClient, authProvider);

            // We need to trigger an initial authentication to ensure we have a token
            await authProvider.AuthenticateRequestAsync(new RequestInformation());

            return new GraphApiService(graphClient);
        }

        public async Task<List<EmailItem>> GetEmailsFromFolderPathAsync(string fullFolderPath, string testSenderEmail, bool? isInTestMode)
        {
            if (string.IsNullOrWhiteSpace(fullFolderPath) || !fullFolderPath.StartsWith("\\\\"))
            {
                throw new ArgumentException("Invalid folder path format. Path must start with '\\\\'.", nameof(fullFolderPath));
            }

            var results = new List<EmailItem>();
            try
            {
                var (_, folderPath) = ParseOutlookPath(fullFolderPath);
                string targetFolderId = await GetFolderIdFromPathAsync(folderPath);
                if (string.IsNullOrEmpty(targetFolderId))
                {
                    throw new System.IO.DirectoryNotFoundException($"The Outlook folder specified by the path '{fullFolderPath}' could not be found.");
                }

                var messages = await _graphClient.Me.MailFolders[targetFolderId].Messages
                                     .GetAsync(requestConfiguration =>
                                     {
                                         requestConfiguration.QueryParameters.Select = new[] { "subject", "receivedDateTime", "sender", "from", "hasAttachments", "internetMessageId" };
                                     });

                foreach (var message in messages.Value)
                {
                    results.Add(new EmailItem
                    {
                        Emailid = message.InternetMessageId,
                        Subject = message.Subject,
                        ReceivedTime = message.ReceivedDateTime?.DateTime ?? DateTime.MinValue,
                        SenderName = message.Sender?.EmailAddress?.Name,
                        SenderEmailAddress = message.Sender?.EmailAddress?.Address,
                        AttachmentCount = message.HasAttachments == true ? 1 : 0
                    });
                }
            }
            catch (Exception ex)
            {
                Log.RecordError($"Failed to retrieve emails using Graph API for path '{fullFolderPath}'.", ex, "GetEmailsFromFolderPathAsync");
                throw;
            }

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

        private async Task<string> GetFolderIdFromPathAsync(string path)
        {
            var folderNames = path.Split(new[] { '\\' }, StringSplitOptions.RemoveEmptyEntries);
            var currentFolders = await _graphClient.Me.MailFolders.GetAsync();
            string currentFolderId = null;

            foreach (var folderName in folderNames)
            {
                var foundFolder = currentFolders.Value.FirstOrDefault(f => f.DisplayName.Equals(folderName, StringComparison.OrdinalIgnoreCase));
                if (foundFolder == null) return null;

                currentFolderId = foundFolder.Id;
                currentFolders = await _graphClient.Me.MailFolders[currentFolderId].ChildFolders.GetAsync();
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
            return (parts[0], parts[1]);
        }
    }

    /// <summary>
    /// This is the correct, modern IAuthenticationProvider implementation.
    /// </summary>
    public class PublicClientAuthenticationProvider : IAuthenticationProvider
    {
        private readonly IPublicClientApplication _clientApp;
        private readonly string[] _scopes;

        public PublicClientAuthenticationProvider(string clientId, string tenantId, string[] scopes)
        {
            _scopes = scopes;
            _clientApp = PublicClientApplicationBuilder.Create(clientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, tenantId)
                .WithDefaultRedirectUri()
                .Build();
        }

        public async Task AuthenticateRequestAsync(RequestInformation request, Dictionary<string, object> additionalAuthenticationContext = null, CancellationToken cancellationToken = default)
        {
            var accounts = await _clientApp.GetAccountsAsync();
            AuthenticationResult authResult;

            try
            {
                authResult = await _clientApp.AcquireTokenSilent(_scopes, accounts.FirstOrDefault()).ExecuteAsync(cancellationToken);
            }
            catch (MsalUiRequiredException)
            {
                authResult = await _clientApp.AcquireTokenInteractive(_scopes).ExecuteAsync(cancellationToken);
            }

            request.Headers.Add("Authorization", new[] { $"Bearer {authResult.AccessToken}" });
        }
    }
}