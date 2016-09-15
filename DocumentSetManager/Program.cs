using System;
using System.Collections.Generic;
using System.Security;
using System.Configuration;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.DocumentSet;
using System.IO;
using System.Linq;
using RapidCircle.SharePoint.DocumentSets;

namespace DocumentSetManager
{
    class Program
    {
        const string ProjectConfigFile = "project.config";

        static void Main(string[] args)
        {

            string webUrl;
            string userName;
            SecureString password;
            Dictionary<string, string> docsets = new Dictionary<string, string>();
            string documentLibraryTitle;
            string documentsFolder;
            bool majorVersionsOnly = false;
            string documentSetName;
            string excludeFolders = null;
            ClientContext cc;


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
                password = securePassword;

                documentLibraryTitle = config.AppSettings.Settings["DocumentLibraryTitle"].Value;
                documentsFolder = config.AppSettings.Settings["DocumentsFolder"].Value;
                documentSetName = config.AppSettings.Settings["DocumentSetName"].Value;
                excludeFolders = config.AppSettings.Settings["ExcludeFolders"].Value.ToLower();

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
                documentLibraryTitle = Console.ReadLine();

                Console.WriteLine("Enter the folder that contains the documents to add the the document set:");
                documentsFolder = Console.ReadLine();

                Console.WriteLine("Enter the name of the document set:");
                documentSetName = Console.ReadLine();

                Console.WriteLine("Enter [1] to copy only Major versions or [2] for all versions:");
                majorVersionsOnly = Console.ReadLine().Equals("1");
            }

            cc = new ClientContext(webUrl);

            cc.Credentials = new SharePointOnlineCredentials(userName, password);
            cc.Load(cc.Web);
            cc.ExecuteQuery();

            //Setup Mappings.
            string[] documentSetMapping = documentSetName.Split(',');
            string[] folderNameMapping = documentsFolder.Split(',');

            for (int i = 0; i < documentSetMapping.Length; i++)
            {
                docsets.Add(folderNameMapping[i], documentSetMapping[i]);
            }

            DocumentSetManagerConfiguration dsmconfig = new DocumentSetManagerConfiguration(documentLibraryTitle, docsets);
            dsmconfig.MajorVersionsOnly = majorVersionsOnly;
            if (excludeFolders != null) dsmconfig.ExcludedFolders = excludeFolders.Split(',').ToList();

            RapidCircle.SharePoint.DocumentSets.DocumentSetManager dsm = new RapidCircle.SharePoint.DocumentSets.DocumentSetManager(cc.Web, dsmconfig);
            dsm.Run();

            cc.Dispose();
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
