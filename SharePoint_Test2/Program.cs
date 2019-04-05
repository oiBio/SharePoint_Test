using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Sharing;

namespace SharePoint_Test2
{
    class Program
    {
        static void Main(string[] args)
        {
            string siteUrl = $"https://aveosolutions.sharepoint.com/";
            string accountEmail = "";
            string password = "";
            using (var context = new ClientContext(siteUrl))
            {
                context.Credentials = new SharePointOnlineCredentials(accountEmail, password.ToSecureString());
                CreateAndShareFolder(context, $"1234 Customername {DateTime.Now.ToString("mm.ss")}", $"TEST1{new Random().Next(100,999)}@oibio.net");
            }
            Console.ReadLine();
        }

        public static void CreateAndShareFolder(ClientContext context, string folderName, string userEmail)
        {
            var listTitle = "Dokumente";
            var rootFolder = "Testfolder_ROOT/";

            var folder = CreateFolder(context.Web, listTitle, rootFolder + folderName);

            var users = new List<UserRoleAssignment>();
            users.Add(new UserRoleAssignment()
            {
                UserId = userEmail,
                Role = Role.Edit,
            });

            var serverRelativeUrl = folder.ServerRelativeUrl;
            var absoluteUrl = new Uri(context.Url).GetLeftPart(UriPartial.Authority) + serverRelativeUrl;


            // Diest konnte eine Spur auf den richtign Weg sein!
            // https://social.technet.microsoft.com/wiki/contents/articles/39365.sharepoint-online-sharing-settings-with-csom.aspx?Sort=MostUseful&PageIndex=1#Property

            //var spoTenant = new Microsoft.Online.SharePoint.TenantAdministration.Tenant(context);
            //context.Load(spoTenant);
            //context.ExecuteQuery();
            //spoTenant.RequireAcceptingAccountMatchInvitedAccount = true;
            //context.Load(spoTenant);
            //context.ExecuteQuery();

            /* User gets Email, but with "public" Sharing-Link */
            //var userSharingResults = DocumentSharingManager.UpdateDocumentSharingInfo(context,
            //    absoluteUrl,
            //    users,
            //    validateExistingPermissions: false,
            //    additiveMode: true,
            //    sendServerManagedNotification: true,
            //    customMessage: null,
            //    includeAnonymousLinksInNotification: true,
            //    propagateAcl: false);

            /* User gets Email, but needs an MS-Account to View */
            var userSharingResults = DocumentSharingManager.UpdateDocumentSharingInfo(context, absoluteUrl, users,
                validateExistingPermissions: false,
                additiveMode: false,
                sendServerManagedNotification: true,
                customMessage: null,
                includeAnonymousLinksInNotification: false,
                propagateAcl: true);

            context.ExecuteQuery();
            if (userSharingResults.FirstOrDefault()?.Status != true)
                Console.WriteLine($"Fehler beim Erstellen {userSharingResults.FirstOrDefault()?.Message}");

        }


        // From Stackoverflow: https://stackoverflow.com/a/22010815/3062062
        public static Folder CreateFolder(Web web, string listTitle, string fullFolderUrl)
        {
            if (string.IsNullOrEmpty(fullFolderUrl))
                throw new ArgumentNullException("fullFolderUrl");
            var list = web.Lists.GetByTitle(listTitle);
            return CreateFolderInternal(web, list.RootFolder, fullFolderUrl);
        }

        private static Folder CreateFolderInternal(Web web, Folder parentFolder, string fullFolderUrl)
        {
            var folderUrls = fullFolderUrl.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string folderUrl = folderUrls[0];
            var curFolder = parentFolder.Folders.Add(folderUrl);
            web.Context.Load(curFolder);
            web.Context.ExecuteQuery();

            if (folderUrls.Length > 1)
            {
                var subFolderUrl = string.Join("/", folderUrls, 1, folderUrls.Length - 1);
                return CreateFolderInternal(web, curFolder, subFolderUrl);
            }
            return curFolder;
        }

    }

    public static class Extensions
    {
        /// <summary>
        /// Returns a Secure string from the source string
        /// </summary>
        /// <param name="Source"></param>
        /// <returns></returns>
        public static SecureString ToSecureString(this string source)
        {
            if (string.IsNullOrWhiteSpace(source))
                return null;
            else
            {
                SecureString result = new SecureString();
                foreach (char c in source.ToCharArray())
                    result.AppendChar(c);
                return result;
            }
        }
    }
  
}
