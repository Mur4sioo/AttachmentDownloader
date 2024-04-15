using System;
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
}
class Program
{
    
    private static UserCredential Login(string googleClientId, string googleClientSecret, string[] scopes)
    {
        ClientSecrets secrets = new ClientSecrets()
        {
            ClientId = googleClientId,
            ClientSecret = googleClientSecret
        };
        return GoogleWebAuthorizationBroker.AuthorizeAsync(secrets, scopes, "user", CancellationToken.None).Result;
    }
    static void Main(string[] args)
    {
        var userId = "me";
        var googleClientId = "";
        var googleClientSecret = "";
        string[] scopes = new[] { Google.Apis.Gmail.v1.GmailService.Scope.GmailReadonly };
        UserCredential credential = Login(googleClientId, googleClientSecret, scopes);
        var initializer = new BaseClientService.Initializer()
        {
            HttpClientInitializer = credential
        };
        using var gmail = new GmailService(initializer);

        var messageIds = GmailExtensions.GetMessages(gmail,userId).Select(message => message.Id);

        foreach(var messageId in messageIds)
        {
            var message = gmail
                .Users
                .Messages
                .Get(userId, messageId)
                .Execute();
            var parts = message.Payload.Parts;
            if (parts != null)
            {
                foreach (var part in parts)
                {
                    var attachmentId = message
                        .Payload
                        .Body
                        .AttachmentId;
                    if (attachmentId != null)
                    {
                        var base64Data = gmail
                            .Users
                            .Messages
                            .Attachments
                            .Get(userId, messageId, attachmentId)
                            .Execute()
                            .Data;

                        // do something with attachment
                        if (base64Data != null)
                        {
                            var fileType = part.MimeType;
                            if (fileType == "application/pdf")
                            {
                                var downloadFile =
                                    Convert.FromBase64String(base64Data.Replace('-', '+').Replace('_', '/'));
                                var fileName = $"{messageId}.pdf";
                                File.WriteAllBytes(fileName, downloadFile);
                                Console.WriteLine($"downloaded PDF file : {fileName}");
                            }
                            else
                                Console.WriteLine("Incorrect format");
                        }
                    }
                }
            }
        }

        Console.ReadKey();
    }
}