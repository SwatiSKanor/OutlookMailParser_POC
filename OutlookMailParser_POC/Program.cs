using Microsoft.Exchange.WebServices.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using static Microsoft.Exchange.WebServices.Data.SearchFilter;


using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System;
using System.Configuration;

namespace OutlookMailParser_POC
{
    public class Program
    {
        private static readonly ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010, TimeZoneInfo.Utc);
        private static ExtendedPropertyDefinition unprocessed { get; set; }

        private readonly string lastReadKey;
        private readonly string accountEmailId;

        //static void Main(string[] args)
        //{
        //    #region MS exchange login 
        //    //       try
        //    //       {
        //    //           ExchangeService service = new ExchangeService(ExchangeVersion.Exchange2010, TimeZoneInfo.Utc);
        //    //           service.Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx");



        //    //          // service.AutodiscoverUrl("TESTFastaffTravelUpdates@fastaff.com");
        //    //        //   service.Credentials = new NetworkCredential("swati.kanor14@outlook.com", "Talentcube@1234", "");

        //    //           service.Credentials = new NetworkCredential("swati@outlook.com", "Talentcube@1234", "");

        //    //           EmailMessage email = new EmailMessage(service);


        //    //           SearchFilter time = new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeReceived, DateTime.UtcNow);



        //    //           ItemView view = new ItemView(100);

        //    //           PropertySet set = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.UniqueBody, ItemSchema.Attachments,
        //    //EmailMessageSchema.Body, EmailMessageSchema.From, EmailMessageSchema.ToRecipients, EmailMessageSchema.DateTimeReceived, EmailMessageSchema.HasAttachments); //ItemSchema.TextBody,

        //    //           FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, time, view).Result; //use call1 if you add subject filer // call1 variable                

        //    //           foreach(Item item in findResults)
        //    //           {
        //    //               var mail = item as EmailMessage;
        //    //              var message = EmailMessage.Bind(service, item.Id, set).Result;

        //    //               var subject = message.Subject;
        //    //               var body = message.Body;
        //    //               var mailId = mail.Id.UniqueId;
        //    //               var from = message.From;

        //    //               if (message.HasAttachments)
        //    //               {
        //    //                   foreach(var att in message.Attachments)
        //    //                   {
        //    //                       if (att is FileAttachment && !att.IsInline) //also probably need to add a check for content type as there could be images as part of attachments
        //    //                       {
        //    //                           var file = (att as FileAttachment);
        //    //                           file.Load().Wait();
        //    //                           if (new List<string> { "pdf", "doc", "docx", "odt", "odtx", "rtf", "txt" }.Contains(file.Name.Split('.').Last().ToLower()))
        //    //                           {


        //    //                           }
        //    //                       }
        //    //                   }
        //    //               }

        //    //           }
        //    //       }
        //    //       catch (Exception ex)
        //    //       {


        //    //           throw;
        //    //       }
        //    #endregion

        //}
        static bool RedirectionCallback(string url)
        {
            // Return true if the URL is an HTTPS URL.
            return url.ToLower().StartsWith("https://");
        }






            static async System.Threading.Tasks.Task Main(string[] args)
            {
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


                

                PropertySet set = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject, ItemSchema.UniqueBody, ItemSchema.Attachments,
             EmailMessageSchema.Body, EmailMessageSchema.From, EmailMessageSchema.ToRecipients, EmailMessageSchema.DateTimeReceived, EmailMessageSchema.HasAttachments); //ItemSchema.TextBody,


                // Make an EWS call
                var folders = service.FindFolders(WellKnownFolderName.MsgFolderRoot, new FolderView(100)).Result;

                SearchFilter time = new SearchFilter.IsLessThanOrEqualTo(ItemSchema.DateTimeReceived, DateTime.UtcNow);
                ItemView view = new ItemView(100);
               FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, time, view).Result; //use call1 if you add subject filer // call1 variable                

                foreach (Item item in findResults)
                {
                    var mail = item as EmailMessage;
                    var message = EmailMessage.Bind(service, item.Id, set).Result;

                    var subject = message.Subject;
                    var body = message.Body;
                    var mailId = mail.Id.UniqueId;
                    var from = message.From;

                    if (message.HasAttachments)
                    {
                        foreach (var att in message.Attachments)
                        {
                            if (att is FileAttachment && !att.IsInline) //also probably need to add a check for content type as there could be images as part of attachments
                            {
                                var file = (att as FileAttachment);
                                file.Load().Wait();
                                if (new List<string> { "pdf", "doc", "docx", "odt", "odtx", "rtf", "txt" }.Contains(file.Name.Split('.').Last().ToLower()))
                                {


                                }
                            }
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
        }
    }
