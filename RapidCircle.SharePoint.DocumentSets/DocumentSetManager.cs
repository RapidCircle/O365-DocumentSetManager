using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client.DocumentSet;
using RapidCircle.SharePoint.DocumentSets.Extensions;

namespace RapidCircle.SharePoint.DocumentSets
{
    public class DocumentSetManager
    {
        private ClientRuntimeContext Ctx { get; set; }
        public Web DocumentWeb { get; set; }
        public DocumentSetManagerConfiguration Config { get; set; }

        public DocumentSetManager(Web documentWeb, DocumentSetManagerConfiguration config)
        {
            this.Ctx = documentWeb.Context;
            DocumentWeb = documentWeb;
            Config = config;
        }

        public void Run()
        {
            List documentLibrary = DocumentWeb.Lists.GetByTitle(Config.DocumentLibrary);
            Folder rootFolder = documentLibrary.RootFolder;
            Ctx.Load(documentLibrary);
            Ctx.Load(rootFolder, i => i.Folders, i=>i.Name);
            Ctx.ExecuteQuery();

            foreach (KeyValuePair<string, string> folderToDocumentMap in Config.FolderToDocumentSetMapping)
            {
                ContentType documentSetCt = DocumentWeb.ContentTypes.GetByName(folderToDocumentMap.Value);
                Folder folder = rootFolder.Folders.GetByName(folderToDocumentMap.Key);
                DocumentSetTemplate dst = DocumentSetTemplate.GetDocumentSetTemplate(Ctx, documentSetCt);
                Ctx.Load(documentSetCt);
                Ctx.Load(folder);
                Ctx.Load(dst, i => i.DefaultDocuments);
                Ctx.ExecuteQuery();

                dst.Clear();

                CamlQuery cq = CreateAllFilesQuery();
                cq.FolderServerRelativeUrl = folder.ServerRelativeUrl;
                ListItemCollection documents = documentLibrary.GetItems(cq);

                Ctx.Load(documents, icol => icol.Include(i => i.File));
                Ctx.Load(documents, icol => icol.Include(i => i.ContentType));
                Ctx.Load(documents, icol => icol.Include(i => i.ContentType.Id));
                Ctx.ExecuteQuery();

                foreach (ListItem document in documents)
                {
                    ContentType docCt = DocumentWeb.ContentTypes.GetByName(document.ContentType.Name);
                    Ctx.ExecuteQuery();

                    string contentTypeFolder = ContentTypeFolder(document, rootFolder, folder);
                    if (Config.ExcludedFolders.Contains(contentTypeFolder.TrimEnd('/'), StringComparer.CurrentCultureIgnoreCase))
                        continue;

                    dst.Add(document, folderToDocumentMap.Value, contentTypeFolder, docCt.Id, Config.MajorVersionsOnly);
                    dst.Update(true);
                    Ctx.ExecuteQuery();
                }

            }
        }

        private static CamlQuery CreateAllFilesQuery()
        {
            var qry = new CamlQuery();
            qry.ViewXml =
                "<View Scope=\"RecursiveAll\"><Query><Where><Eq><FieldRef Name=\"FSObjType\" /><Value Type=\"Integer\">0</Value></Eq></Where></Query></View>";
            return qry;
        }

        private string ContentTypeFolder(ListItem document, Folder rootFolder, Folder contentTypeFolder)
        {
            char[] splitter = {'/'};
            string[] path = document.File.ServerRelativeUrl.Split(splitter);
            string folder = path[path.GetLength(0) - 2];

            if (folder == rootFolder.Name || folder == contentTypeFolder.Name)
                folder = "";
            else
                folder = folder + @"/";

            return folder;
        }

    }
}
