using System;
using System.Collections;
using System.Linq;
using System.Net;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
namespace AttachmentsDownloader;
public static class GmailExtensions
{
    public static IList<Message> GetMessages(this GmailService service, string userId)
        => service.Users.Messages.List(userId).Execute().Messages;
    public static async Task<IList<Message>> GetMessagesAsync(this GmailService service, string userId)
        => (await service.Users.Messages.List(userId).ExecuteAsync()).Messages;
}

class Program
{
    private static async Task<GmailService> CreateGmailService(string googleClientId, string googleClientSecret)
    {
        UserCredential credential = await Login(googleClientId, googleClientSecret, GmailService.Scope.GmailReadonly);
        var initializer = new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential
        };
        return new GmailService(initializer);
    }
    
    private static Task<UserCredential> Login(string googleClientId, string googleClientSecret, params string[] scopes)
    {
        ClientSecrets secrets = new ClientSecrets()
        {
            ClientId = googleClientId,
            ClientSecret = googleClientSecret
        };
        return GoogleWebAuthorizationBroker.AuthorizeAsync(secrets, scopes, "user", CancellationToken.None);
    }

    
    private static async Task Main(string[] args)
    {
        var userId = "me";
        using var gmail = await CreateGmailService(googleClientId: "", googleClientSecret: "");
        var messageIds = (await gmail.GetMessagesAsync(userId)).Select(message => message.Id);
        var options = new ParallelOptions()
        {
            MaxDegreeOfParallelism = 10,
        };
        await Parallel.ForEachAsync(messageIds, options, async (messageId, cancellationToken) =>
        {
            var message = await gmail
                .Users
                .Messages
                .Get(userId, messageId)
                .ExecuteAsync(cancellationToken);
            var parts = message.Payload.Parts;
            if (parts != null)
            {
                foreach (var part in parts)
                {
                    await GetAttachment(message, gmail, userId, part);
                }
            }
        });

        Console.ReadKey();
    }
    private static async Task SaveAttachment(MessagePart part, string base64Data, string messageId)
    {
        var fileType = part.MimeType;
        if (fileType == "application/pdf")
        {
            var downloadFile =
                Convert.FromBase64String(base64Data.Replace('-', '+').Replace('_', '/'));
            var fileName = $"{messageId}.pdf";
            await File.WriteAllBytesAsync(fileName, downloadFile);
            Console.WriteLine($"downloaded PDF file : {fileName}");
        }
        else
            Console.WriteLine("Incorrect format");
    }

    private static async Task GetAttachment(GmailService gmail, string userId, string messageId, string attachmentId,
        MessagePart part)
    {
        var base64Data = (await gmail
            .Users
            .Messages
            .Attachments
            .Get(userId, messageId, attachmentId)
            .ExecuteAsync())
            .Data;

        // do something with attachment
        if (base64Data != null)
        {
            await SaveAttachment(part, base64Data, messageId);
        }
    }

    private static async Task GetAttachment(Message message, GmailService gmail, string userId, MessagePart part)
    {
        var attachmentId = message
            .Payload
            .Body
            .AttachmentId;
        if (attachmentId != null)
        {
            await GetAttachment(gmail, userId, message.Id, attachmentId, part);
        }
    }
}