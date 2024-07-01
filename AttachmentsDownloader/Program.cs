using System;
using System.Collections;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
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
    {
        var request = service.Users.Messages.List(userId);
        request.MaxResults = 500;
        var messages = (await request.ExecuteAsync()).Messages;
        return messages;
    }
    
}

class Program
{
    private static HashSet<string> fileChecksums = new HashSet<string>();
    [Obsolete("Obsolete")]
    private static async Task<GmailService> CreateGmailService()
    {

        UserCredential credential;
        await using (var stream = new FileStream("client_secrets.json", FileMode.Open, FileAccess.Read))
        {
            const string credPath = "token.json";
            credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                GoogleClientSecrets.Load(stream).Secrets,
                new[] { GmailService.Scope.GmailReadonly },
                "user",
                CancellationToken.None,
                new FileDataStore(credPath, true));
            test
        }
        
        return new GmailService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = "Gmail Attachment Downloader"
        
        });
}

    private record AttachmentInfo(string AttachmentId, string MessageId, string MimeType, string FileName);
    private static IEnumerable<AttachmentInfo> GetAllAttachmentIds(MessagePart part, string messageId)
    {
        var queue = new Queue<MessagePart>();
        queue.Enqueue(part);
        while (queue.TryDequeue(out var item))
        {
            if (item.Body.AttachmentId is not null)
            {
                yield return new( item.Body.AttachmentId, messageId, item.MimeType, item.Filename);
            }

            if (item.Parts is null)
            {
                continue;
            }
            foreach (var subpart in item.Parts)
                queue.Enqueue(subpart);
        }
    }
    
    private static bool IsAttachmentPdf(AttachmentInfo attachment)
    {
        if (attachment.MimeType.Equals("application/octet-stream", StringComparison.OrdinalIgnoreCase)
            && (attachment.FileName.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)))
            return true;
        if (attachment.MimeType == "application/pdf")
        {
            return true;
        }

        return false;
    }

    [Obsolete("Obsolete")]
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

            if (IsAttachmentPdf(attachment))
            {
                var title = message.Payload.Headers.FirstOrDefault(header => header.Name == "Subject");
                Console.WriteLine(title.Value);
                await SaveAttachment(gmail, userId, messageId, attachment);
            }
            else
                Console.WriteLine("Incorrect format.");
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
        var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "download", "TEMP_FILE.pdf");
        DownloadTempFile(attachment);
        string fileChecksum =
            ComputeMd5Checksum(path);
        if (fileChecksums.Contains(fileChecksum))
        {
            Console.WriteLine("File was already downloaded.");
            File.Delete(path);
        }
        else
        {
            fileChecksums.Add(fileChecksum);
            Console.WriteLine("New file downloading...");
            File.Delete(path);
            DownloadNewFile(gmail, userId, messageId, attachment);
        }
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
            var match = Regex.Match(subject, "BIEDRONKA\\s+(\\d+)");
            detailName = match.Groups[1].Value;
            subject = $"B{detailName}";
        }
        else if(subject.Contains("HEBE"))
        {
            var match = Regex.Match(subject, "HEBE\\s+R(\\d+)");
            detailName = match.Groups[1].Value;
            subject = $"R{detailName}";
        }
        else
        {
            subject = "DEAFULT";
        }
        return subject;
    }

    private static void DownloadTempFile(MessagePartBody attachment)
    {
        var downloadFile = Convert.FromBase64String((attachment.Data).Replace('-', '+').Replace('_', '/'));
        var fileName = "TEMP_FILE.pdf";
        var downloadDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "download", fileName);
        File.WriteAllBytes(downloadDirectory,downloadFile);
    }
    private static void DownloadNewFile(GmailService gmail, string userId, string messageId, MessagePartBody attachment)
    {
        var downloadFile = Convert.FromBase64String((attachment.Data).Replace('-', '+').Replace('_', '/'));
        var fileNamePart = GetTitleText(gmail, userId, messageId);
        var fileName = $"{fileNamePart}_PN_{DateTime.Today:yyyMMdd}.pdf";
        var downloadDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "download", fileName);
        File.WriteAllBytes(downloadDirectory,downloadFile);
        Console.WriteLine($"downloaded file {fileName}");
    }
    private static string ComputeMd5Checksum(string filePath)
    {
        using var md5 = MD5.Create();
        using var stream = File.OpenRead(filePath);
        var hash = md5.ComputeHash(stream);
        return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
    }
}