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
using Newtonsoft.Json;

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
    private static List<string> fileNames = new List<string>();

    private static void loadHashes()
    {
        string filehashset = "hashset.json";
        if (File.Exists(filehashset))
        {
            var jsonDataHash = File.ReadAllText(filehashset);
            fileChecksums = JsonConvert.DeserializeObject<HashSet<string>>(jsonDataHash) ?? new HashSet<string>();
        }
        else
        {
            File.Create("hashset.json");
        }
        string fileName = "filename.json";
        if (File.Exists(fileName))
        {
            var jsonDataName = File.ReadAllText(fileName);
            fileNames = JsonConvert.DeserializeObject<List<string>>(jsonDataName) ?? new List<string>();
        }
        else
        {
            File.Create("filename.json");
        }
    }

    private static void saveHashes()
    {
        string fileHashset = "hashset.json";
        string jsonDataHash = JsonConvert.SerializeObject(fileChecksums, Formatting.Indented);
        File.WriteAllText(fileHashset, jsonDataHash);

        string fileName = "filename.json";
        string jsonDataName = JsonConvert.SerializeObject(fileNames, Formatting.Indented);
        File.WriteAllText(fileName, jsonDataName);
    }
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
        loadHashes();
        await Parallel.ForEachAsync(messageIds, options, async (messageId, cancellationToken) =>
        {
            await GetAttachmentsFromMessageId(gmail, userId, messageId, cancellationToken);
        });
        saveHashes();
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
        var title = $"{GetTitleText(gmail, userId, messageId)}";
        var date = $"{GetDateText(gmail, userId, messageId)}";
        var filename = $"{GetFileName(title, date)}.pdf";
        var path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "download", filename);
        var downloadFile = Convert.FromBase64String((attachment.Data).Replace('-', '+').Replace('_', '/'));
        if (!IsDownloaded(downloadFile))
        {
            Console.WriteLine($"Downloading ... {filename}");
            File.WriteAllBytes(path, downloadFile);
        }
        else
        {
            Console.WriteLine($"File {filename} already downloaded.");
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

    private static string GetDateText(GmailService gmail, string userId, string messageId)
    {
        var message = gmail.Users.Messages.Get(userId, messageId).Execute();
        var dateHeader = message.Payload.Headers?.FirstOrDefault(header => header.Name.Equals("Date", StringComparison.OrdinalIgnoreCase));
        var dateString = dateHeader?.Value ?? "";
        dateString.ToString("yyyyMMdd");
        return dateString;
    }

    private static string GetFileName(string title, string date)
    {
        var name = $"{title}_PN_{date}";
        int i = 1;
        while (fileNames.Contains(name))
        {
            name = $"{title}_PN_Z{i}_{date}";
            i++;
        }
        fileNames.Add(name);
        return name;
    }
    
    private static bool IsDownloaded(byte[] attachment)
    {
        var hash = ComputeHash(attachment);
        if (fileChecksums.Contains(hash))
        {
            return true;
        }
        fileChecksums.Add(hash);
        return false;
    }
    private static string ComputeHash(byte[] byteArray)
    {
        using (SHA256 sha256 = SHA256.Create())
        {
            byte[] hashBytes = sha256.ComputeHash(byteArray);
            return BitConverter.ToString(hashBytes).Replace("-", "").ToLower();
        }
    }
}