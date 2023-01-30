/*##########################################################################################################
# Disclaimer
# This sample code, scripts, and other resources are not supported under any Microsoft standard support 
# program or service and are meant for illustrative purposes only.
#
# The sample code, scripts, and resources are provided AS IS without warranty of any kind. Microsoft 
# further disclaims all implied warranties including, without limitation, any implied warranties of 
# merchantability or of fitness for a particular purpose. The entire risk arising out of the use or 
# performance of this material and documentation remains with you. In no event shall Microsoft, its 
# authors, or anyone else involved in the creation, production, or delivery of the sample be liable 
# for any damages whatsoever (including, without limitation, damages for loss of business profits, 
# business interruption, loss of business information, or other pecuniary loss) arising out of the 
# use of or inability to use the samples or documentation, even if Microsoft has been advised of 
# the possibility of such damages.
##########################################################################################################*/
using Microsoft.Exchange.WebServices.Data;
using Microsoft.Identity.Client;
using System;
using System.Configuration;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace SearchFolderApp
{
    class Program
    {
        static async System.Threading.Tasks.Task Main(string[] args)
        {
            /* Register an application with Azure Active Directory that has Exchange full_access_as_app permission.
             * https://outlook.office365.com/full_access_as_app 
             *
             * Enter app registration settings in the App.config file.
             */

            IConfidentialClientApplication app = ConfidentialClientApplicationBuilder.Create(ConfigurationManager.AppSettings["AppId"] as string)
                .WithTenantId(ConfigurationManager.AppSettings["TenantId"])
                .WithClientSecret(ConfigurationManager.AppSettings["AppSecret"])
                .Build();
            AuthenticationResult auth = await app.AcquireTokenForClient(new string[] { "https://outlook.office365.com/.default" }).ExecuteAsync();

            var service = new ExchangeService();
            service.HttpHeaders.Add("Authorization", $"Bearer {auth.AccessToken}");
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, "Mark8ProjectTeam@M365x612691.onmicrosoft.com");
            service.Url = new Uri("https://outlook.office365.com/ews/exchange.asmx");
            service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, ConfigurationManager.AppSettings["impersonatedUser"] as string);

            GetMailboxSize(service);
            SyncFolderHierarchy(service);
            CreateSearchFolder(service);
            GetArchiverootItems(service);
            GetSearchFolders(service);
            GetMessagesAllProperties(service);
            GetMessagesReducedProperties(service);
            GetFoldersInFolder(service);
            GetMessages(service);
            GetMessagesQuery(service);
            GetChangesByBatchSize(service);
            GetCalendarSelectProperties(service);
            GetArchiveFolder(service);

            Console.WriteLine("Press any key ...");
            Console.ReadKey();

        }

        private static void GetFoldersInFolder(ExchangeService service)
        {
            FolderView view = new FolderView(100);
            var results = service.FindFolders(new FolderId("AAMkADljNDI4ZDI4LTlhOWEtNDdjYy05NGJlLWI1NmNmMzhiYmQ1MwAuAAAAAAD8/brfsm6RRax0jDhbq9egAQCE/MAqGzi1SYwF8BP/ts40AAAAAAEhAAA="), view);
        }

        private static void GetArchiveFolder(ExchangeService service)
        {
            Folder archiveFolderRoot = Folder.Bind(service, WellKnownFolderName.ArchiveMsgFolderRoot);
        }

        private static void GetChangesByBatchSize(ExchangeService service)
        {
            int offset = 0;
            int pageSize = 50;
            bool more = true;

            ItemView view = new ItemView(pageSize, offset, OffsetBasePoint.Beginning);

            PropertySet propSet = (BasePropertySet.IdOnly);
            propSet.Add(ItemSchema.Size);

            view.PropertySet = propSet;

            FindItemsResults<Item> findResults;
            List<EmailMessage> foundemails = new List<EmailMessage>();

            while (more)
            {
                findResults = service.FindItems(WellKnownFolderName.Inbox, view);
                foundemails.AddRange(findResults.Cast<EmailMessage>());
                more = findResults.MoreAvailable;
                if (more)
                {
                    view.Offset += pageSize;
                }
            }

            PropertySet properties = (BasePropertySet.FirstClassProperties);

            more = true;
            int batchSize = 0;
            int maxSize = 1000000;

            List<EmailMessage> emails = new List<EmailMessage>();
            List<EmailMessage> retrievedEmails = new List<EmailMessage>();

            for (int idx = 0; idx < foundemails.Count; idx++)
            {
                if (((batchSize + foundemails[idx].Size) > maxSize) && emails.Count() > 0)
                {
                    service.LoadPropertiesForItems(emails, properties);
                    retrievedEmails.AddRange(emails);
                    emails.Clear();
                    batchSize = 0;
                }
                else
                {
                    emails.Add(foundemails[idx]);
                    batchSize += foundemails[idx].Size;
                }
            }
        }

        private static void GetMessages(ExchangeService service)
        {
            int offset = 0;
            int pageSize = 25;
            bool more = true;
            ItemView view = new ItemView(pageSize, offset, OffsetBasePoint.Beginning);

            view.PropertySet = PropertySet.FirstClassProperties;
            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, view);
        }

        private static void GetMessagesQuery(ExchangeService service)
        {
            int offset = 0;
            int pageSize = 25;
            bool more = true;
            ItemView view = new ItemView(pageSize, offset, OffsetBasePoint.Beginning);
            view.PropertySet = PropertySet.FirstClassProperties;
            SearchFilter.IsEqualTo subjectFilter = new SearchFilter.IsEqualTo(EmailMessageSchema.Subject, "Azure AD Identity Protection Weekly Digest");
            FindItemsResults<Item> findResults = service.FindItems(WellKnownFolderName.Inbox, subjectFilter, view);
        }

        static void GetMailboxSize(ExchangeService service)
        {
            FolderView view = new FolderView(10000);
            view.Traversal = FolderTraversal.Shallow;
            var msgSizeExtended = new ExtendedPropertyDefinition(3592, MapiPropertyType.Long);
            List<ExtendedPropertyDefinition> viewProps = new List<ExtendedPropertyDefinition>
            {
                msgSizeExtended
            };
            view.PropertySet = new PropertySet(viewProps);
            var results = service.FindFolders(WellKnownFolderName.MsgFolderRoot, view);

            Int64 totalSize = 0;
            foreach(var folder in results)
            {
                Int64 folderSize = 0;
                if (folder.TryGetProperty(msgSizeExtended, out folderSize))
                {
                    totalSize += (Int64)folderSize;
                }
            }

            Console.WriteLine($"Total Size: {totalSize}");
        }

        #region Data access methods
        static void GetSearchFolders(ExchangeService service)
        {
            FolderView folderView = new FolderView(1000);
            var folders = service.FindFolders(WellKnownFolderName.SearchFolders, folderView);
            foreach (Folder folder in folders)
            {
                if (folder is SearchFolder)
                {
                    try
                    {
                        SearchFolder searchFolder = SearchFolder.Bind(service, folder.Id);
                        Console.WriteLine(searchFolder.DisplayName);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Unable to get folder {folder.Id.UniqueId}.");
                    }
                }
            }
        }

        static void GetArchiverootItems(ExchangeService service)
        {
            var folders = service.SyncFolderHierarchy(WellKnownFolderName.ArchiveMsgFolderRoot, PropertySet.IdOnly, null);
            foreach (FolderChange folderChange in folders)
            {
                Folder folder = Folder.Bind(service, folderChange.FolderId);
                Console.WriteLine(folder.DisplayName);
                Console.WriteLine(folder.FolderClass);
            }
        }

        static void CreateSearchFolder(ExchangeService service)
        {
            Console.WriteLine("Enter Search Folder Name:");
            var folderName = Console.ReadLine();
            SearchFolder newSearchfolder = new SearchFolder(service);
            newSearchfolder.DisplayName = folderName;
            EmailAddress manager = new EmailAddress("sadie@contoso.com");
            SearchFilter.IsEqualTo fromManagerFilter = new SearchFilter.IsEqualTo(EmailMessageSchema.Sender, manager);
            newSearchfolder.SearchParameters.SearchFilter = fromManagerFilter;
            newSearchfolder.SearchParameters.RootFolderIds.Add(WellKnownFolderName.Inbox);
            newSearchfolder.SearchParameters.Traversal = SearchFolderTraversal.Shallow;
            newSearchfolder.Save(WellKnownFolderName.SearchFolders);
        }

        static void SyncFolderHierarchy(ExchangeService service)
        {
            var propertySet = new PropertySet();
            propertySet.BasePropertySet = BasePropertySet.IdOnly;
            propertySet.Add(FolderSchema.ParentFolderId);
            propertySet.Add(FolderSchema.DisplayName);
            var folderChanges = service.SyncFolderHierarchy(WellKnownFolderName.MsgFolderRoot, propertySet, null);
            foreach (var change in folderChanges)
            {
                Console.WriteLine(change.Folder.DisplayName);
            }
        }

        static void GetMessagesAllProperties(ExchangeService service)
        {
            int offset = 0;
            int pageSize = 50;
            bool more = true;
            ItemView view = new ItemView(pageSize, offset, OffsetBasePoint.Beginning);

            view.PropertySet = PropertySet.IdOnly;
            FindItemsResults<Item> findResults;
            List<EmailMessage> emails = new List<EmailMessage>();

            while (more)
            {
                findResults = service.FindItems(WellKnownFolderName.Inbox, view);
                foreach (var item in findResults.Items)
                {
                    emails.Add((EmailMessage)item);
                }
                more = findResults.MoreAvailable;
                if (more)
                {
                    view.Offset += pageSize;
                }
            }
            PropertySet properties = (BasePropertySet.FirstClassProperties); //A PropertySet with the explicit properties you want goes here
            service.LoadPropertiesForItems(emails, properties);
        }

        static void GetMessagesReducedProperties(ExchangeService service)
        {
            int offset = 0;
            int pageSize = 50;
            bool more = true;
            ItemView view = new ItemView(pageSize, offset, OffsetBasePoint.Beginning);

            view.PropertySet = PropertySet.IdOnly;
            FindItemsResults<Item> findResults;
            List<EmailMessage> emails = new List<EmailMessage>();

            while (more)
            {
                findResults = service.FindItems(new FolderId("AAMkADljNDI4ZDI4LTlhOWEtNDdjYy05NGJlLWI1NmNmMzhiYmQ1MwAuAAAAAAD8/brfsm6RRax0jDhbq9egAQCE/MAqGzi1SYwF8BP/ts40AAAAAAEhAAA="), view);
                foreach (var item in findResults.Items)
                {
                    emails.Add((EmailMessage)item);
                }
                more = findResults.MoreAvailable;
                if (more)
                {
                    view.Offset += pageSize;
                }
            }
            PropertySet properties = (BasePropertySet.IdOnly);
            properties.Add(PostItemSchema.ParentFolderId);
            properties.Add(PostItemSchema.Subject);
            properties.Add(PostItemSchema.HasAttachments);
            properties.Add(PostItemSchema.DateTimeReceived);
            properties.Add(PostItemSchema.From);
            properties.Add(PostItemSchema.Sender);
            properties.Add(PostItemSchema.DisplayTo);
            properties.Add(EmailMessageSchema.Importance);
            properties.Add(EmailMessageSchema.ToRecipients);
            properties.Add(TaskSchema.StartDate);
            service.LoadPropertiesForItems(emails, properties);
        }

        static void GetCalendarSelectProperties(ExchangeService service)
        {
            CalendarFolder calendar = CalendarFolder.Bind(service, WellKnownFolderName.Calendar);

            PropertySet propSet = new PropertySet(BasePropertySet.FirstClassProperties);
            propSet.Add(MeetingRequestSchema.Resources);
            ItemView view = new ItemView(100);
            view.OrderBy.Add(MeetingRequestSchema.Start, SortDirection.Ascending);
            view.PropertySet = propSet;

            CalendarView calendarView = new CalendarView(DateTime.Now, DateTime.Now.AddDays(1));
            var appointments = calendar.FindAppointments(calendarView);
            foreach (Appointment a in appointments)
            {
                Console.Write("Subject: " + a.Subject.ToString() + " ");
                Console.Write("Start: " + a.Start.ToString() + " ");
                Console.Write("End: " + a.End.ToString());
                Console.WriteLine();
            }
        }

        #endregion
    }
}
