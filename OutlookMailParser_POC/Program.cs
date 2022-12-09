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
                .Create("802023a0-57d0-4949-bc02-7c209bcb02e9")  //appId
                .WithClientSecret("OSs8Q~xE1j_crnwshg_ztlyycqewuBPvQh-wEcRP")   //client SECRETE
                .WithTenantId("e5f294ba-7871-4659-a1b4-ba9cbd7c2eed")   //  CLIENT TENANT ID
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
                        new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "rpotnis@psgglobalsolutions.com");

                //Include x-anchormailbox header
             //   service.HttpHeaders.Add("X-AnchorMailbox", "swati@thinkbridge.in");


                //Microsoft.Exchange.WebServices.Data.Folder rootfolder = Microsoft.Exchange.WebServices.Data.Folder.Bind(service, WellKnownFolderName.MsgFolderRoot).Result;
                //rootfolder.Load();
                //foreach( Folder fo in rootfolder)
                //{
                //    var f = fo.
                //}

                     ReplytoSharedInboxEmail();

                ExtendedPropertyDefinition replyEmail = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings, "replyemail", MapiPropertyType.Boolean);


                PropertySet set = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.UniqueBody, ItemSchema.Attachments,
             EmailMessageSchema.Body, EmailMessageSchema.From, EmailMessageSchema.ToRecipients, EmailMessageSchema.DateTimeReceived, EmailMessageSchema.HasAttachments, ItemSchema.InReplyTo, replyEmail); //ItemSchema.TextBody,


                // Make an EWS call
                //SearchFilter foldername = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "swati k");
                //  SearchFilter subjectFilter = new SearchFilter.ContainsSubstring(ItemSchema.Subject, "473616-Test7 T");


            
                SearchFilter time = new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeReceived, DateTime.UtcNow);
                SearchFilter reply = new SearchFilter.Exists(replyEmail);

                time = new SearchFilterCollection(LogicalOperator.Or, time, reply);

                ItemView view = new ItemView(100);

                var findResults = service.FindItems(new FolderId(WellKnownFolderName.Inbox, "psgsharedinbox@psgglobal.com"), time, view).Result;

                //   var findResults1 = service.FindItems(new FolderId(WellKnownFolderName.Root, "compassa@thinkbridge.in"), time, view).Result;

                //    Microsoft.Exchange.WebServices.Data.Folder rootfolder = Microsoft.Exchange.WebServices.Data.Folder.Bind(service, WellKnownFolderName.MsgFolderRoot).Result;

                #region archival movement of mails
                FolderView fview = new FolderView(1000);
                fview.PropertySet = new PropertySet(BasePropertySet.IdOnly);
                fview.PropertySet.Add(FolderSchema.DisplayName);
                fview.PropertySet.Add(FolderSchema.FolderClass);
                fview.PropertySet.Add(FolderSchema.ParentFolderId);

                fview.Traversal = FolderTraversal.Deep;


                var ss = service.FindFolders(new FolderId(WellKnownFolderName.Root, "psgsharedinbox@psgglobal.com"), fview).Result;

                var fss = ss.FirstOrDefault(x => x.DisplayName == "Shared Inbox ArchivalMail");

                if (fss == null)
                {// Create a custom folder.
                    Microsoft.Exchange.WebServices.Data.Folder folder = new Microsoft.Exchange.WebServices.Data.Folder(service);

                    folder.DisplayName = "Shared Inbox ArchivalMail";
                    folder.FolderClass = "IPF.Note";
                    // Save the folder as a child folder of the Inbox.
                    await folder.Save(new FolderId(WellKnownFolderName.MsgFolderRoot, "psgsharedinbox@psgglobal.com"));
                    fss = folder;
                }
                var fclass = fss.FolderClass;


                #endregion

                foreach (Item item in findResults)
                {
                    bool isReplyEmail = false;
                    var mail = item as EmailMessage;
                    var message = EmailMessage.Bind(service, item.Id, set).Result;
                    object responseEmail;
                    message.TryGetProperty(replyEmail, out responseEmail);
                    if (responseEmail != null)
                    {
                        isReplyEmail = (Boolean)responseEmail;

                    }

                    if (!isReplyEmail || message.From.Address != "psgsharedinbox@psgglobal.com")
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
                    else
                    {
                        // Request conversation items. This results in a call to the service.
                        ConversationResponse response = service.GetConversationItems
       (item.ConversationId, set, null, null,
       ConversationSortOrder.TreeOrderAscending).Result;
                        foreach(ConversationNode node in response.ConversationNodes)
                        {
                            foreach (Item nodeItem in node.Items)
                            {
                               await nodeItem.Move(fss.Id);
                            }
                                //await message.Move(fss.Id);

                        }

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

            var mailId = "AAMkADk1MGRmZjg2LTNiNGEtNDc0NS1iZjJlLThkNTZiMDYxNDBiYQBGAAAAAADuPM+QeU3mT7pKjnOTG3yVBwDqkvPFbnQ5R6WkEZICnT8cAAAAAAEMAADqkvPFbnQ5R6WkEZICnT8cAAAHZvZ6AAA=";
            PropertySet set = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.UniqueBody, ItemSchema.Attachments,
         EmailMessageSchema.Body, EmailMessageSchema.From, EmailMessageSchema.ToRecipients, EmailMessageSchema.DateTimeReceived, EmailMessageSchema.HasAttachments, EmailMessageSchema.ReplyTo, EmailMessageSchema.Sender); //ItemSchema.TextBody,

            FolderView fview = new FolderView(1000);
            fview.PropertySet = new PropertySet(BasePropertySet.IdOnly);
            fview.PropertySet.Add(FolderSchema.DisplayName);
            fview.PropertySet.Add(FolderSchema.FolderClass);
            fview.PropertySet.Add(FolderSchema.ParentFolderId);
            fview.PropertySet.Add(FolderSchema.WellKnownFolderName);

            fview.Traversal = FolderTraversal.Deep;


            var ss = service.FindFolders(new FolderId(WellKnownFolderName.Root, "psgsharedinbox@psgglobal.com"), fview).Result;

           // var fss = ss.FirstOrDefault(x => x.DisplayName == "Shared Inbox ArchivalMail");
            var saveFss = ss.FirstOrDefault(x => x.WellKnownFolderName.HasValue && x.WellKnownFolderName.Value == WellKnownFolderName.Inbox);


            EmailMessage msg = (EmailMessage)Item.Bind(service, new ItemId(mailId), set).Result;
            var subject = msg.Subject;

            ExtendedPropertyDefinition replyEmail = new ExtendedPropertyDefinition(DefaultExtendedPropertySet.PublicStrings, "replyemail", MapiPropertyType.Boolean);
          //  msg.From = new EmailAddress { Address = "compassa@thinkbridge.in", Name = "Compass A" };
         //  var r= (EmailMessage) msg.Update(ConflictResolutionMode.NeverOverwrite).Result;
            ResponseMessage responseMessage =  msg.CreateReply(true);
            string myReply = "testing knock! knock!";
            ;
            
            responseMessage.BodyPrefix = myReply;

            // // Send the response message.
            // // This method call results in a CreateItem call to EWS.
            try
            {
                var e = await responseMessage.Save();
               var s=  e.Update(ConflictResolutionMode.AlwaysOverwrite).Result;
                e.From = new EmailAddress { Address = "psgsharedinbox@psgglobal.com" };
                e.SetExtendedProperty(replyEmail, true);
                await e.Send();

                //to set replyemail prop true in reply mail so it will exclude from creating the task in next iterations
                //var replyMessage = responseMessage.Save(saveFss.Id).Result;
                //var replyUpdate = replyMessage.Update(ConflictResolutionMode.AlwaysOverwrite).Result;
                //replyMessage.SetExtendedProperty(replyEmail, true);

                //replyMessage.From = new EmailAddress { Address = "compassa@thinkbridge.in", Name = "Compass A" };
                //await replyMessage.SendAndSaveCopy();
                //  var movedEMail = msg.Move(fss.Id);
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

                var folderPath = context.Web.GetFolderByServerRelativeUrl("Shared Documents/RPA/Shared Inbox POC/" + today + "/TaskId_" + 1111.ToString());
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
                catch (ServerException ex)
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
