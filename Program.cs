using Microsoft.Graph;
using Azure.Identity;
using Microsoft.Identity.Client;
using System;
using System.Threading.Tasks;
using System.Linq;
using System.IO;
using System.Security.Cryptography;
using System.Collections.Generic;
using Microsoft.Kiota.Abstractions.Authentication;
using Azure.Core;

class Program
{
    static async Task<int> Main(string[] args)
    {
        var clientId = Environment.GetEnvironmentVariable("CLIENT_ID");
        var tenantId = Environment.GetEnvironmentVariable("TENANT_ID");

        if (string.IsNullOrEmpty(clientId) || string.IsNullOrEmpty(tenantId))
        {
            Console.WriteLine("CLIENT_ID or TENANT_ID environment variables are not set.");
            return 3;
        }

        var argsOk = ParseArgs(args, out var mailboxAddress, out var folderName, out var messageCount);
        if (!argsOk) return 1;

        var graphClient = await GetAuthenticatedGraphClient(clientId, tenantId);
        if (graphClient == null)
        {
            Console.WriteLine("Authentication failed.");
            return 1;
        }

        var folderId = await GetFolderIdByNameAsync(graphClient, mailboxAddress, folderName);
        if (folderId == null) return 2;

        var messages = await GetMessagesAsync(graphClient, mailboxAddress, folderId, messageCount);
        if (messages != null)
        {
            Console.WriteLine($"Found {messages.Count} messages in '{folderName}' folder.");
            Console.WriteLine("------------------------------");
            foreach (var message in messages)
            {
                Console.WriteLine($"FROM: {message.Sender?.EmailAddress?.Address}");

                var toRecipients = message.ToRecipients;
                if (toRecipients != null)
                {
                    var toRecipientsAsString = string.Join("; ", toRecipients.Select(r => r.EmailAddress?.Address));
                    Console.WriteLine($"TO: {toRecipientsAsString}");
                }

                var ccRecipients = message.CcRecipients;
                if (ccRecipients != null)
                {
                    var ccRecipientsAsString = string.Join("; ", ccRecipients.Select(r => r.EmailAddress?.Address));
                    Console.WriteLine($"CC: {ccRecipientsAsString}");
                }

                Console.WriteLine($"Subject: {message.Subject}");

                var body = message.Body?.Content;
                if (!string.IsNullOrEmpty(body))
                {
                    body = body.Substring(0, Math.Min(100, body.Length));
                    body = body.Replace("\n", "\\n").Replace("\r", "\\r");
                    Console.WriteLine($"Body: {body}");
                }
                else
                {
                    Console.WriteLine("Body: (empty)");
                }

                Console.WriteLine("------------------------------");
            }
        }

        return 0;
    }

    private static bool ParseArgs(string[] args, out string mailboxAddress, out string folderName, out int messageCount)
    {
        mailboxAddress = "me";
        folderName = "Inbox";
        messageCount = 10;

        for (int i = 0; i < args.Length; i++)
        {
            if (args[i] == "--mailbox" && i + 1 < args.Length)
            {
                mailboxAddress = args[i + 1];
            }
            else if (args[i] == "--folder" && i + 1 < args.Length)
            {
                folderName = args[i + 1];
            }
            else if (args[i] == "--messages" && i + 1 < args.Length)
            {
                if (int.TryParse(args[i + 1], out int count))
                {
                    messageCount = count;
                }
            }
            else if (args[i].StartsWith("--"))
            {
                Console.WriteLine("Invalid argument: " + args[i]);
                return false;
            }
        }

        return true;
    }

    private static async Task<GraphServiceClient?> GetAuthenticatedGraphClient(string clientId, string tenantId)
    {
        var options = new DeviceCodeCredentialOptions
        {
            AuthorityHost = AzureAuthorityHosts.AzurePublicCloud,
            ClientId = clientId,
            TenantId = tenantId,
            DeviceCodeCallback = (code, cancellation) =>
            {
                Console.WriteLine(code.Message);
                return Task.FromResult(0);
            },
            TokenCachePersistenceOptions = new TokenCachePersistenceOptions() { Name = "email-reader-cache" }
        };

        var TokenFp = "token.json";
        DeviceCodeCredential deviceCodeCredential;
        if (File.Exists(TokenFp))
        {
            using var fileStream = new FileStream(TokenFp, FileMode.Open, FileAccess.Read);
            options.AuthenticationRecord = await AuthenticationRecord.DeserializeAsync(fileStream);
            deviceCodeCredential = new DeviceCodeCredential(options);
        }
        else
        {
            deviceCodeCredential = new DeviceCodeCredential(options);
            var authenticationRecord = await deviceCodeCredential.AuthenticateAsync(new TokenRequestContext(scopes));
            using var fileStream1 = new FileStream(TokenFp, FileMode.Create, FileAccess.Write);
            await authenticationRecord.SerializeAsync(fileStream1);
        }

        return new GraphServiceClient(deviceCodeCredential, scopes);
    }

    private static async Task<string?> GetFolderIdByNameAsync(GraphServiceClient graphClient, string mailboxAddress, string folderName)
    {
        var specialFolderName = AdjustSpecialFolderNames(folderName);
        var response = mailboxAddress == "me"
            ? await graphClient.Me.MailFolders.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Select = new string[] { "id", "displayName" };
                requestConfiguration.QueryParameters.Top = 100;
            })
            : await graphClient.Users[mailboxAddress].MailFolders.GetAsync((requestConfiguration) =>
            {
                requestConfiguration.QueryParameters.Select = new string[] { "id", "displayName" };
                requestConfiguration.QueryParameters.Top = 100;
            });

        var mailFolders = response?.Value;
        if (mailFolders != null)
        {
            var folder = mailFolders.FirstOrDefault(f => f.DisplayName != null && f.DisplayName.Equals(folderName, StringComparison.OrdinalIgnoreCase));
            if (!string.IsNullOrEmpty(folder?.Id))
            {
                return folder.Id;
            }
        }

        Console.WriteLine($"Folder '{folderName}' not found.");
        return null;
    }

    private static string? AdjustSpecialFolderNames(string folderName)
    {
        return folderName.ToLower() switch
        {
            "inbox" => "inbox",
            "sent items" => "sentitems",
            "drafts" => "drafts",
            "deleted items" => "deleteditems",
            "outbox" => "outbox",
            "junk email" => "junkemail",
            "archive" => "archive",
            _ => null,
        };
    }

    private static async Task<List<Microsoft.Graph.Models.Message>?> GetMessagesAsync(GraphServiceClient graphClient, string mailboxAddress, string folderId, int messageCount)
    {
        var response = mailboxAddress == "me"
            ? await graphClient.Me.MailFolders[folderId].Messages.GetAsync((requestConfiguration) =>
                InitGetMessagesQueryParameters(requestConfiguration.QueryParameters, messageCount))
            : await graphClient.Users[mailboxAddress].MailFolders[folderId].Messages.GetAsync((requestConfiguration) =>
                InitGetMessagesQueryParameters(requestConfiguration.QueryParameters, messageCount));
        return response?.Value;
    }

    private static void InitGetMessagesQueryParameters(Microsoft.Graph.Me.MailFolders.Item.Messages.MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters parameters, int messageCount)
    {
        parameters.Select = new string[] { "sender", "toRecipients", "ccRecipients", "subject", "body" };
        parameters.Orderby = new string[] { "receivedDateTime desc" };
        parameters.Top = messageCount;
    }

    private static void InitGetMessagesQueryParameters(Microsoft.Graph.Users.Item.MailFolders.Item.Messages.MessagesRequestBuilder.MessagesRequestBuilderGetQueryParameters parameters, int messageCount)
    {
        parameters.Select = new string[] { "sender", "toRecipients", "ccRecipients", "subject", "body" };
        parameters.Orderby = new string[] { "receivedDateTime desc" };
        parameters.Top = messageCount;
    }

    private static string[] scopes = { "User.Read", "Mail.Read" };
}
