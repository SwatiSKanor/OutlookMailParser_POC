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
using System.Text;
using System.Globalization;

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

                ExtendedPropertyDefinition replyEmail = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings, "replyemail", MapiPropertyType.Boolean);


                PropertySet set = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.UniqueBody, ItemSchema.Attachments,
             EmailMessageSchema.Body, EmailMessageSchema.From, EmailMessageSchema.ToRecipients, EmailMessageSchema.DateTimeReceived, EmailMessageSchema.HasAttachments,ItemSchema.InReplyTo, replyEmail); //ItemSchema.TextBody,


                // Make an EWS call
                //    SearchFilter foldername = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "FastaffTravelUpdates");
                //  SearchFilter subjectFilter = new SearchFilter.ContainsSubstring(ItemSchema.Subject, "[EXTERNAL] License nursys Jessica Welch (369580)");
                //    var folders = service.FindFolders(WellKnownFolderName.MsgFolderRoot, foldername, new FolderView(100)).Result;

                SearchFilter time = new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeReceived, DateTime.UtcNow);
                SearchFilter reply = new SearchFilter.Exists(replyEmail);

                   time = new SearchFilterCollection(LogicalOperator.Or, time, reply);

                ItemView view = new ItemView(100);

                var findResults = service.FindItems(new FolderId(WellKnownFolderName.Inbox, "compassa@thinkbridge.in"),time,view).Result;

                foreach (Item item in findResults)
                {
                    bool isReplyEmail = false;
                    var mail = item as EmailMessage;
                    var message = EmailMessage.Bind(service, item.Id, set).Result;
                    object responseEmail;
                    message.TryGetProperty(replyEmail, out responseEmail);
                    if (responseEmail != null)
                    {
                      isReplyEmail   = (Boolean)responseEmail;

                    }
                    if (!isReplyEmail)
                    {
                        var subject = message.Subject;
                        var recipient = message.ToRecipients;
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

                                var nodeConvItem_Subject = nodeItem.Subject;

                                if (nodeItem.HasAttachments)
                                {
                                    foreach (Microsoft.Exchange.WebServices.Data.Attachment attach in nodeItem.Attachments)
                                    {
                                        FileAttachment fileAttachment = attach as FileAttachment;
                                        if (fileAttachment != null)
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
                                                    ContentStream = memoryStream
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

                        var mailId = "AAMkADcwYmE1ZjFlLTFkNGMtNGUzMC04ZTA5LWU5YWY2MmIzYTI0MABGAAAAAAB+kXU0TcU4QILxCi62HT45BwBZ3Hlh4f4VR6AUTaq0jP5TAAAAAAEMAABZ3Hlh4f4VR6AUTaq0jP5TAAAHEdOSAAA=";
            PropertySet set = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.UniqueBody, ItemSchema.Attachments,
         EmailMessageSchema.Body, EmailMessageSchema.From, EmailMessageSchema.ToRecipients, EmailMessageSchema.DateTimeReceived, EmailMessageSchema.HasAttachments,EmailMessageSchema.ReplyTo,EmailMessageSchema.Sender); //ItemSchema.TextBody,

            EmailMessage msg = (EmailMessage)Item.Bind(service, new ItemId(mailId), set).Result;
            var subject = msg.Subject;

            //  msg.ToRecipients.Add("abhishek@thinkbridge.in");
            //  msg.Body = new MessageBody("test sample.");
            //ExtendedPropertyDefinition PidTagReplyRecipientEntries = new ExtendedPropertyDefinition(0x004F, MapiPropertyType.Binary);
            //ExtendedPropertyDefinition PidTagReplyRecipientNames = new ExtendedPropertyDefinition(0x0050, MapiPropertyType.String);
            // msg.SetExtendedProperty(PidTagReplyRecipientEntries, ConvertHexStringToByteArray(GenerateFlatList("swati@thinkbridge.in", "Swati K")));
            //msg.InReplyTo = "replymail";
            //msg.SetExtendedProperty(PidTagReplyRecipientNames, "Swati K");

            ExtendedPropertyDefinition replyEmail =  new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings, "replyemail", MapiPropertyType.Boolean);
          //  msg.SetExtendedProperty(replyEmail,true);
            //  await msg.Reply(new MessageBody { BodyType = BodyType.Text, Text = "This sample text on date 4th OCT. Test 4. 20:46" },true);

            ResponseMessage responseMessage = msg.CreateReply(true);
            msg.From = new EmailAddress { Address = "swati@thinkbridge.in", Name = "Swati K" };
            string myReply = "This is the message body of the email reply.21:20";
            responseMessage.BodyPrefix = myReply;
        
            // // Send the response message.
            // // This method call results in a CreateItem call to EWS.
            try
            {
                // to set replyemail prop true in reply mail so it will exclude from creating the task in next iterations
                var  e =  await responseMessage.Save();
                e.SetExtendedProperty(replyEmail, true);
                await e.Send();
            

            }
            catch (System.Exception ex)
            {

            }

        }



  

    internal static String GenerateFlatList(String SMTPAddress, String DisplayName)
        {
            String abCount = "01000000";
            String AddressId = GenerateOneOff(SMTPAddress, DisplayName);
            return abCount + BitConverter.ToString(INT2LE((AddressId.Length / 2) + 4)).Replace("-", "") + BitConverter.ToString(INT2LE(AddressId.Length / 2)).Replace("-", "") + AddressId;
        }

        internal static String GenerateOneOff(String SMTPAddress, String DisplayName)
        {
            String Flags = "00000000";
            String ProviderUid = "812B1FA4BEA310199D6E00DD010F5402";
            String Version = "0000";
            String xFlags = "0190";
            String DisplayNameHex = BitConverter.ToString(UnicodeEncoding.Unicode.GetBytes(DisplayName + "\0")).Replace("-", "");
            String SMTPAddressHex = BitConverter.ToString(UnicodeEncoding.Unicode.GetBytes(SMTPAddress + "\0")).Replace("-", "");
            String AddressType = BitConverter.ToString(UnicodeEncoding.Unicode.GetBytes("SMTP" + "\0")).Replace("-", "");
            return Flags + ProviderUid + Version + xFlags + DisplayNameHex + AddressType + SMTPAddressHex;
        }
        internal static byte[] INT2LE(int data)
        {
            byte[] b = new byte[4];
            b[0] = (byte)data;
            b[1] = (byte)(((uint)data >> 8) & 0xFF);
            b[2] = (byte)(((uint)data >> 16) & 0xFF);
            b[3] = (byte)(((uint)data >> 24) & 0xFF);
            return b;
        }
        internal static byte[] ConvertHexStringToByteArray(string hexString)
        {
            if (hexString.Length % 2 != 0)
            {
                throw new ArgumentException(String.Format(CultureInfo.InvariantCulture, "The binary key cannot have an odd number of digits: {0}", hexString));
            }

            byte[] HexAsBytes = new byte[hexString.Length / 2];
            for (int index = 0; index < HexAsBytes.Length; index++)
            {
                string byteValue = hexString.Substring(index * 2, 2);
                HexAsBytes[index] = byte.Parse(byteValue, NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            }

            return HexAsBytes;

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
