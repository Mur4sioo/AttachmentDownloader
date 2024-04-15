using System;
using System.Linq;
using Google.Apis.Gmail.v1;
using Google.Apis.Gmail.v1.Data;
using Google.Apis.Services;
namespace AttachmentsDownloader;

class Program
{
    
    static void Main(string[] args)
    {
        var userId = "kacpermurawski2000@gmail.com";
        
        var initializer = new BaseClientService.Initializer();
        // set properties on intializer
        using var gmail = new GmailService(initializer);

        var messageIds = gmail
            .Users
            .Messages
            .List(userId)
            .Execute()
            .Select(message => message.Id);

        foreach(var messageId in messageIds)
        {
            var message = gmail
                .Users
                .Message
                .Get(userId, messageId)
                .Execute();
            var attachmentId = message
                .Payload
                .Body
                .AttachmentId;
            if(attachmentId is null)
                continue;

            var base64Data = gmail
                .Users
                .Messages
                .Attachments
                .Get(userId, messageId, attachmentId)
                .Execute()
                .Data;

            // do something with attachment
        }
    }
}