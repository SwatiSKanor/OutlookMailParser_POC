using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using static Microsoft.Exchange.WebServices.Data.SearchFilter;
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System;
using System.Configuration;
using System.IO;
using Microsoft.WindowsAzure.Storage.Blob;
using Microsoft.WindowsAzure.Storage;
using Microsoft.SharePoint.Client;
using System.Security;
using CSOMDemo;
using System.Net;
using AuthenticationManager = CSOMDemo.AuthenticationManager;
using System.Threading.Tasks;

namespace OutlookMailParser_POC
{
    public class Program
    {
        private static readonly ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2015, TimeZoneInfo.Utc);
        public const string ConnectionString = "DefaultEndpointsProtocol=https;AccountName=adminaccinternalstorage;AccountKey=LSACTbnz5p57zc4NXkQEE+UKxa2C5WrdgwMhWlt2ir+NeayrS8hyDfXKmWiZIZ/6X1yJwVzH28LOGbs6BoGQwA==;";


        static async System.Threading.Tasks.Task Main(string[] args)
            {
          //  CreateNewFolderSharePoint();
            // Using Microsoft.Identity.Client 4.22.0  
            // thinkbridge active directory - Shared inbox outlook poc 
            var cca = ConfidentialClientApplicationBuilder
                .Create("ed92b7ff-019d-4e70-b04c-dd38a95fdeec")  //appId
                .WithClientSecret("qPu8Q~rRu5AEwMy9U62tCxW87bl9OO1PMI~zwcWk")   //client SECRETE
                .WithTenantId("b69d82df-4ebe-474d-9ac7-00efbf13427e")   //  CLIENT TENANT ID
                .Build();


            var ewsScopes = new string[] { "https://outlook.office365.com/.default" };

                try
                {
                    var authResult = await cca.AcquireTokenForClient(ewsScopes).ExecuteAsync();

                    // Configure the ExchangeService with the access token
                //    var ewsClient = new ExchangeService();
                    service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");
                service.Credentials = new OAuthCredentials(authResult.AccessToken);
                service.ImpersonatedUserId =
                        new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "swati@thinkbridge.in");

                //Include x-anchormailbox header
                service.HttpHeaders.Add("X-AnchorMailbox", "swati@thinkbridge.in");

                EmailMessage email = new EmailMessage(service);

                ReplytoSharedInboxEmail();


                PropertySet set = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.UniqueBody, ItemSchema.Attachments,
             EmailMessageSchema.Body, EmailMessageSchema.From, EmailMessageSchema.ToRecipients, EmailMessageSchema.DateTimeReceived, EmailMessageSchema.HasAttachments); //ItemSchema.TextBody,


                // Make an EWS call
                SearchFilter foldername = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "FastaffTravelUpdates");

            //    var folders = service.FindFolders(WellKnownFolderName.MsgFolderRoot, foldername, new FolderView(100)).Result;
                SearchFilter time = new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeReceived, DateTime.UtcNow);
              //  SearchFilter subjectFilter = new SearchFilter.ContainsSubstring(ItemSchema.Subject, "[EXTERNAL] License nursys Jessica Welch (369580)");
                ItemView view = new ItemView(100);

                //var findResults = service.FindItems(WellKnownFolderName.Inbox,time, view).Result;

                // shared inbox mail access

                FolderId SharedMailbox = new FolderId(WellKnownFolderName.Inbox, "compassa@thinkbridge.in");
                ItemView itemView = new ItemView(100);
                FindItemsResults<Item> findResults = service.FindItems(SharedMailbox, "", itemView).Result;

                //  var findResults = folders.Folders.FirstOrDefault().FindItems(time, view).Result;
                foreach (Item item in findResults)
                {
                    var mail = item as EmailMessage;
                    var message = EmailMessage.Bind(service, item.Id, set).Result;

                    var subject = message.Subject;
                    var body = message.Body;
                    var mailId = mail.Id.UniqueId;
                    var from = message.From;
                    var conversationId = item.ConversationId;
                   
                   
                    // Request conversation items. This results in a call to the service.
                    ConversationResponse response = service.GetConversationItems
   (conversationId, set, null, null,
   ConversationSortOrder.TreeOrderDescending).Result;


                    foreach (ConversationNode node in response.ConversationNodes)
                    {
                        foreach (Item nodeItem in node.Items)
                        {
                            var nodeConvItemBody = nodeItem.UniqueBody;

                            var nodeConvItem_Subject= nodeItem.Subject;
                        
                            if (nodeItem.HasAttachments)
                            {
                                foreach (Microsoft.Exchange.WebServices.Data.Attachment attach in nodeItem.Attachments)
                                {
                                    FileAttachment fileAttachment = attach as FileAttachment;
                                    if(fileAttachment != null)
                                    {
                                        fileAttachment.Load().Wait();
                                        if (new List<string> { "pdf", "doc", "docx", "odt", "odtx", "rtf", "txt", "jpeg", "png", "jpg" }.Contains(fileAttachment.Name.Split('.').Last().ToLower()))
                                        {
                                            var memoryStream = new System.IO.MemoryStream(fileAttachment.Content);
                                            memoryStream.Seek(0, SeekOrigin.Begin);
                                            memoryStream.ToArray();
                                            var fileCreationInfo = new FileCreationInformation
                                            {
                                                //Content = fileAttachment.Content.ToArray(),
                                                Overwrite = true,
                                                Url = fileAttachment.Name,
                                                ContentStream =memoryStream
                                            };
                                           // var img = SaveFileToBlobStorage(fileAttachment.Name, fileAttachment.ContentType, fileAttachment.Content, "", "");
                                            var url = UploaToSharepoint(fileCreationInfo, mailId);

                                        }
                                    }
                                 
                                }
                            }
                        }
                    }
                  
                  //  await message.Reply( "Reply for test.",false);
                    #region this code belongs to attachments not received in chains.
                    //if (message.HasAttachments)
                    //{
                    //    foreach (var att in message.Attachments)
                    //    {
                    //        if (att is FileAttachment) //also probably need to add a check for content type as there could be images as part of attachments
                    //        {
                    //            var file = (att as FileAttachment);
                    //            file.Load().Wait();
                    //            if (new List<string> { "pdf", "doc", "docx", "odt", "odtx", "rtf", "txt","jpeg","png","jpg" }.Contains(file.Name.Split('.').Last().ToLower()))
                    //            {
                    //               var img =  SaveFileToBlobStorage(file.Name, file.ContentType, file.Content, "", "");

                    //            }
                    //        }
                    //    }
                    //}
                    #endregion
                }

            }
            catch (MsalException ex)
                {
                    Console.WriteLine($"Error acquiring access token: {ex}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error: {ex}");
                }

                if (System.Diagnostics.Debugger.IsAttached)
                {
                    Console.WriteLine("Hit any key to exit...");
                    Console.ReadKey();
                }
            }
        public static async void ReplytoSharedInboxEmail()
        {
            var mailId = "AQMkADcwYmE1ZjFlLTFkNGMtNGUzMC04ZTA5LWU5YWY2MmIzYTI0MABGAAADfpF1NE3FOECC8Qouth0+OQcAWdx5YeH+FUegFE2qtIz+UwAAAgEMAAAAWdx5YeH+FUegFE2qtIz+UwAAAgVbAAAA";
            PropertySet set = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.UniqueBody, ItemSchema.Attachments,
         EmailMessageSchema.Body, EmailMessageSchema.From, EmailMessageSchema.ToRecipients, EmailMessageSchema.DateTimeReceived, EmailMessageSchema.HasAttachments); //ItemSchema.TextBody,

            EmailMessage msg = (EmailMessage)Item.Bind(service, new ItemId(mailId), set).Result;
            var subject = msg.Subject;
            var rec = msg.From;
            ResponseMessage responseMessage = msg.CreateReply(false);
        
            string myReply = "This is the message body of the email reply.";
            responseMessage.BodyPrefix = myReply;
           
            responseMessage.IsDeliveryReceiptRequested = true;
            responseMessage.ToRecipients.Add("priyanka@thinkbridge.in");

            // Send the response message.
            // This method call results in a CreateItem call to EWS.
            try
            {
                await responseMessage.SendAndSaveCopy();

            }
            catch (Exception ex)
            {

            }

        }

        public static string SaveFileToBlobStorage(string fileName, string contentType, byte[] content, string storageContainer, string contentDisposition = "")
        {
            var storageAccount = CloudStorageAccount.Parse(ConnectionString);
            CloudBlobClient blobclient = storageAccount.CreateCloudBlobClient();

            //var serviceProperties = blobclient.GetServicePropertiesAsync().Result;
            //serviceProperties.DefaultServiceVersion = "2015-04-05";
            //blobclient.SetServicePropertiesAsync(serviceProperties);

            CloudBlobContainer container;
            container = blobclient.GetContainerReference("image");
            if (container.CreateIfNotExistsAsync().Result)
                container.SetPermissionsAsync(new BlobContainerPermissions() { PublicAccess = BlobContainerPublicAccessType.Blob });

            try
            {
                CloudBlockBlob clientBlob = container.GetBlockBlobReference(fileName);
                clientBlob.Properties.ContentType = contentType;
                if (contentDisposition != "")
                {
                    clientBlob.Properties.ContentDisposition = contentDisposition;
                }
                else
                {
                    clientBlob.Properties.ContentDisposition = "attachment;filename=" + fileName;
                }
                Stream byteStream = new MemoryStream(content);
                clientBlob.UploadFromStreamAsync(byteStream);
                return clientBlob.StorageUri.PrimaryUri.AbsoluteUri;
            }
            catch (Exception ex)
            {
            }

            return string.Empty;
        }

        public static string UploaToSharepoint(FileCreationInformation fileCreationInfo, string msgId)
        {

            Uri site = new Uri("https://psgglobal.sharepoint.com/sites/RPA_AdminAccounts");
            string password = "1Jaiho!Jaiho!";
            string user = "rpa@psgglobal.com";
            SecureString secureString = new SecureString();
            password.ToList().ForEach(secureString.AppendChar);
            using (AuthenticationManager authenticationManager = new AuthenticationManager())
            using (var context = authenticationManager.GetContext(site, user, secureString))
            {

                var folderPath = context.Web.GetFolderByServerRelativeUrl("Shared Documents/RPA/Shared Inbox POC");
                context.Load(folderPath, k => k.Name, k => k.Files, k => k.Folders);
                context.ExecuteQuery();
                Microsoft.SharePoint.Client.File uploadFile = folderPath.Files.Add(fileCreationInfo);


                context.Load(uploadFile);
                context.ExecuteQuery();
                var documentUrl = site.Host + uploadFile.ServerRelativeUrl;

             
                return documentUrl;
                
            }
             
            return string.Empty;
        }

        public static void CreateNewFolderSharePoint()
        {

            Uri site = new Uri("https://psgglobal.sharepoint.com/sites/RPA_AdminAccounts");
            string password = "1Jaiho!Jaiho!";
            string user = "rpa@psgglobal.com";
            SecureString secureString = new SecureString();
            password.ToList().ForEach(secureString.AppendChar);
            using (AuthenticationManager authenticationManager = new AuthenticationManager())
            using (var context = authenticationManager.GetContext(site, user, secureString))
            {

                var today = DateTime.UtcNow.Date.ToString("MM/dd/yyyy");

                var folderPath = context.Web.GetFolderByServerRelativeUrl("Shared Documents/RPA/Shared Inbox POC/"+today+"/TaskId_" +1111.ToString() );
                try
                {
                    context.Load(folderPath, k => k.Name, k => k.Files, k => k.Folders);
                    context.ExecuteQuery();
                    if (!folderPath.Exists)
                    {

                        var taskFolder = folderPath.Folders.Add("TaskId_" + 111.ToString());
                        context.Load(taskFolder);
                        context.ExecuteQuery();
                        Console.WriteLine("New folder Created");
                    }
                    else
                    {
                        context.Load(folderPath, k => k.Name, k => k.Files, k => k.Folders);
                        Console.WriteLine("Folder already exists");

                    }
                }
                catch(ServerException ex)
                {
                    var folderPath1 = context.Web.GetFolderByServerRelativeUrl("Shared Documents/RPA/Shared Inbox POC/" + today + "/");

                    var taskFolder = folderPath1.Folders.Add("TaskId_" + 111.ToString());
                    context.Load(taskFolder);
                    context.ExecuteQuery();
                    Console.WriteLine("New folder Created");

                }
              

                //var newF =  folderPath.Folders.Add(today);
                //context.Load(newF);
                //context.ExecuteQuery();

            }
              
        }
    }
    }
