using System;
using System.Security;
using System.Configuration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.DocumentSet;
using System.IO;
using System.Linq;

namespace DocumentSetManager
{
    class Program
    {
        static ClientContext cc;

        private static string docFolder;
        private static string docSetContentTypeId;
        private static string docSetName;

        const string ProjectConfigFile = "project.config";
        private static string listTitle;
        private static string excludeFodlersContaining;

        private static List dl;

        static void Main(string[] args)
        {

            string webUrl;
            string userName;
            SecureString password;

            if (System.IO.File.Exists(ProjectConfigFile))
            {
                ExeConfigurationFileMap configMap = new ExeConfigurationFileMap();
                configMap.ExeConfigFilename = @"project.config";
                Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None);

                SecureString securePassword = new SecureString();
                foreach (char letter in config.AppSettings.Settings["Password"].Value)
                    securePassword.AppendChar(letter);

                webUrl = config.AppSettings.Settings["WebURL"].Value;
                userName = config.AppSettings.Settings["UserName"].Value;

                docFolder = config.AppSettings.Settings["DocumentsFolder"].Value;
                docSetContentTypeId = config.AppSettings.Settings["DocumentSetContentID"].Value;
                docSetName = config.AppSettings.Settings["DocumentSetName"].Value;
                listTitle = config.AppSettings.Settings["ListTitle"].Value;
                excludeFodlersContaining = config.AppSettings.Settings["ExcludeFoldersContaining"].Value.ToLower();

                password = securePassword;
            }
            else
            {
                Console.WriteLine("Enter the URL of the SharePoint Online site hosting your projects:");
                webUrl = Console.ReadLine();

                Console.WriteLine("Enter your user name:");
                userName = Console.ReadLine();

                Console.WriteLine("Enter your password.");
                password = GetPasswordFromConsoleInput();

                Console.WriteLine("Enter the title of the library containing the documents:");
                excludeFodlersContaining = Console.ReadLine();

                Console.WriteLine("Enter the folder that contains the documents to add the the document set:");
                docFolder = Console.ReadLine();

                Console.WriteLine("Enter the document set ID:");
                docSetContentTypeId = Console.ReadLine();

                Console.WriteLine("Enter the name of the document set:");
                docSetName = Console.ReadLine();
            }

            cc = new ClientContext(webUrl);

            cc.Credentials = new SharePointOnlineCredentials(userName, password);
            cc.Load(cc.Web);
            cc.ExecuteQuery();

            //Document Set
            ContentType dsCt = GetDocumentSet(docSetContentTypeId);
            //Delivery Document
            //TODO: This needs to come from the document being added. 
            ContentType ddCt = GetDocumentSet("0x010100617E4FF37CCCF3448B647C453333DD0500DD5195652DF7F4449F14DC04FECE3F5E");

            DocumentSetTemplate dst = DocumentSetTemplate.GetDocumentSetTemplate(cc, dsCt);
            dl = cc.Web.Lists.GetByTitle(listTitle);
            cc.Load(ddCt);
            cc.Load(dst.DefaultDocuments);
            cc.Load(dl.RootFolder);
            cc.ExecuteQuery();

            DeleteAllExistingDocuments(dst);

            ListItemCollection documents = dl.GetItems(CreateAllFilesQuery());
            cc.Load(documents, icol => icol.Include(i => i.File));
            cc.Load(documents, icol => icol.Include(i => i.ContentType));
            cc.ExecuteQuery();

            foreach (ListItem li in documents)
            {
                char[] splitter = { '/' };
                string[] path = li.File.ServerRelativeUrl.Split(splitter);
                string folder = path[path.GetLength(0) - 2];

                if (folder == dl.RootFolder.Name || folder == docFolder)
                    folder = "";
                else
                {
                    folder = folder + @"/";

                    string[] exclusions = excludeFodlersContaining.Split(',');
                    if (exclusions.Any(folder.ToLower().Contains))
                        continue;
                }

                ClientResult<Stream> data = li.File.OpenBinaryStream();
                cc.Load(li.File);
                cc.ExecuteQuery();

                if (data != null)
                {
                    string myPage = "Placeholder";
                    MemoryStream repo = new MemoryStream(System.Text.Encoding.UTF8.GetBytes(myPage));
                    dst.DefaultDocuments.Add(folder + li.File.Name, ddCt.Id, repo.ToArray());
                    dst.Update(true);
                    cc.ExecuteQuery();
                    repo.Close();

                    MemoryStream memStream = new MemoryStream();
                    data.Value.CopyTo(memStream);
                    memStream.Seek(0, SeekOrigin.Begin);
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(cc, cc.Web.ServerRelativeUrl + "/_cts/" + docSetName + "/" + li.File.Name, memStream, true);
                    memStream.Close();
                }
            }
            dst.Update(true);
            cc.ExecuteQuery();
            cc.Dispose();
        }

        private static void DeleteAllExistingDocuments(DocumentSetTemplate dst)
        {
            //Delete Documents.
            string[] docs = new string[200];
            int j = 0;
            foreach (var document in dst.DefaultDocuments)
            {
                docs[j] = document.Name;
                j++;
            }

            foreach (var document in docs)
            {
                if (string.IsNullOrEmpty(document)) break;
                dst.DefaultDocuments.Remove(document);
                dst.Update(true);
                cc.ExecuteQuery();
            }
        }

        public static CamlQuery CreateAllFilesQuery()
        {
            var qry = new CamlQuery();
            qry.FolderServerRelativeUrl = dl.RootFolder.ServerRelativeUrl + "/" + docFolder;

            qry.ViewXml = "<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name=\"FSObjType\" /><Value Type=\"Integer\">0</Value></Eq></Where></Query></View>";

            return qry;
        }

        private static ContentType GetDocumentSet(string contentTypeId)
        {
            ContentType ct = cc.Web.ContentTypes.GetById(contentTypeId);
            cc.ExecuteQuery();

            if (ct != null)
                return ct;

            return null;
        }

        private static SecureString GetPasswordFromConsoleInput()
        {
            ConsoleKeyInfo info;

            //Get the user's password as a SecureString
            SecureString securePassword = new SecureString();
            do
            {
                info = Console.ReadKey(true);
                if (info.Key != ConsoleKey.Enter)
                {
                    securePassword.AppendChar(info.KeyChar);
                }
            }
            while (info.Key != ConsoleKey.Enter);
            return securePassword;
        }
    }
}
