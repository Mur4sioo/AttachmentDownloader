using System;
using System.Collections;
using System.Linq;
using System.Net;
using System.Text.RegularExpressions;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;

namespace AttachmentsDownloader;
public static class GmailExtensions
{
    public static async Task<IList<Message>> GetMessagesAsync(this GmailService service, string userId)
        => (await service.Users.Messages.List(userId).ExecuteAsync()).Messages;
}

class Program
{
    private static async Task<GmailService> CreateGmailService()
    {

        UserCredential credential;
        using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
        {
            string credPath = "token.json";
            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.Load(stream).Secrets,
                new[] { GmailService.Scope.GmailReadonly },
                "user",
                CancellationToken.None,
                new FileDataStore(credPath, true)).Result;
        }
        
        return new GmailService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = "Gmail Attachment Downloader"
        
        });
}

    private record AttachmentInfo(string AttachmentId, string MessageId, string MimeType);
    private static IEnumerable<AttachmentInfo> GetAllAttachmentIds(MessagePart part, string messageId)
    {
        var queue = new Queue<MessagePart>();
        queue.Enqueue(part);
        while (queue.TryDequeue(out var item))
        {
            if (item.Body.AttachmentId is not null)
            {
                yield return new( item.Body.AttachmentId, messageId, item.MimeType);
            }

            if (item.Parts is null)
            {
                continue;
            }
            foreach (var subpart in item.Parts)
                queue.Enqueue(subpart);
        }
    }

    private static async Task Main(string[] args)
    {
        using var gmail = await CreateGmailService();
        var userId = "me";
        var messageIds = (await gmail.GetMessagesAsync(userId)).Select(message => message.Id);
        var options = new ParallelOptions()
        {
            MaxDegreeOfParallelism = 1,
        };
        await Parallel.ForEachAsync(messageIds, options, async (messageId, cancellationToken) =>
        {
            await GetAttachmentsFromMessageId(gmail, userId, messageId, cancellationToken);
        });

        
    }

    private static async Task GetAttachmentsFromMessageId(GmailService gmail, string userId, string messageId,
        CancellationToken cancellationToken)
    {
        var message = await gmail
            .Users
            .Messages
            .Get(userId, messageId)
            .ExecuteAsync(cancellationToken);
        if (message.Payload is null)
        {
            return;
        }

        foreach (var attachment in GetAllAttachmentIds(message.Payload, messageId))
        {
            if (attachment.MimeType != "application/pdf")
            {
                Console.WriteLine("Incorrect format");
            }
            else
            {
                await SaveAttachment(gmail, userId, messageId, attachment);
            }
        }
    }
    
    private static async Task SaveAttachment(GmailService gmail, string userId, string messageId, AttachmentInfo attachmentInfo)
    {
        var attachment = await gmail
            .Users
            .Messages
            .Attachments
            .Get(userId, attachmentInfo.MessageId, attachmentInfo.AttachmentId)
            .ExecuteAsync();
        
        
        var subject = GetTitleText(gmail, userId, messageId);
        var base64Date = attachment.Data;
        var downloadFile = Convert.FromBase64String(base64Date.Replace('-', '+').Replace('_', '/'));
        var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var downloadDirectory = Path.Combine(baseDirectory, "download");
        var fileName = $"{subject}_{DateTime.Today:yyyMMdd}.pdf";
        var fullPath = Path.Combine(downloadDirectory, fileName);
        await File.WriteAllBytesAsync(fullPath, downloadFile);
        Console.WriteLine($"downloaded PDF file : {fileName}");
    }

    private static string GetTitleText(GmailService gmail, string userId, string messageId)
    {
        var tempTitleText = gmail.Users.Messages.Get(userId, messageId);
        tempTitleText.Format = UsersResource.MessagesResource.GetRequest.FormatEnum.Metadata;
        tempTitleText.MetadataHeaders = new List<string>{"Subject"};
        var titleText = tempTitleText.Execute();
        var subjectHeader = titleText.Payload.Headers.FirstOrDefault(header => header.Name == "Subject");
        string subject = subjectHeader?.Value?? "Deafult_filename";
        string detailName;
        if (subject.Contains("BIEDRONKA"))
        {
            Match match = Regex.Match(subject, "BIEDRONKA\\s+(\\d+)");
            detailName = match.Groups[1].Value;
            subject = $"B{detailName}";
        }
        else if(subject.Contains("HEBE"))
        {
            Match match = Regex.Match(subject, "HEBE R(\\d+)");
            detailName = match.Groups[1].Value;
            subject = $"R{detailName}";
        }
        else
        {
            subject = "DEAFULT";
        }
        return subject;
    }
}