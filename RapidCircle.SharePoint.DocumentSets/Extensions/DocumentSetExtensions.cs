using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.DocumentSet;

namespace RapidCircle.SharePoint.DocumentSets.Extensions
{
    public static class DocumentSetExtensions
    {
        public static void Clear (this DocumentSetTemplate documentSet)
        {
            var ctx = documentSet.Context;

            string[] docs = new string[200];
            int j = 0;
            foreach (var document in documentSet.DefaultDocuments)
            {
                docs[j] = document.Name;
                j++;
            }

            foreach (var document in docs)
            {
                if (string.IsNullOrEmpty(document)) break;
                documentSet.DefaultDocuments.Remove(document);
                documentSet.Update(true);
                ctx.ExecuteQuery();
            }
        }

        public static void Add(this DocumentSetTemplate documentSet, ListItem documentItem, string documentSetName, string folder, ContentTypeId docId, bool majorVersionOnly)
        {
            var ctx = documentSet.Context;
            Uri webUrl = new Uri(ctx.Url);

            Microsoft.SharePoint.Client.File documentFile = documentItem.File;
            ctx.Load(documentFile);
            ctx.Load(documentFile.Versions);
            ctx.ExecuteQuery();

            ClientResult<Stream> documentStream = null;
            if (majorVersionOnly)
            {
                if (documentFile.Versions.Count.Equals(0)) return;

                string majorVersionLabel = documentFile.MajorVersion + ".0";
                foreach (FileVersion fileVersion in documentItem.File.Versions)
                {
                    if (fileVersion.VersionLabel.Equals(majorVersionLabel))
                    {
                        documentStream = fileVersion.OpenBinaryStream();
                        ctx.ExecuteQuery();
                        break;
                    }
                }
            }
            else
                documentStream = documentItem.File.OpenBinaryStream();

            //Use a place holder to workaround filesize limitation with the DefaultDocuments API. 
            //Place holder is inserted via the API, and then overwritten later. 
            if (documentStream != null)
            {
                string placeholderPage = "Placeholder";
                MemoryStream repo = new MemoryStream(Encoding.UTF8.GetBytes(placeholderPage));
                documentSet.DefaultDocuments.Add(folder + documentFile.Name, docId, repo.ToArray());
                documentSet.Update(true);
                ctx.ExecuteQuery();
                repo.Close();

                MemoryStream memStream = new MemoryStream();
                documentStream.Value.CopyTo(memStream);
                memStream.Seek(0, SeekOrigin.Begin);
                Microsoft.SharePoint.Client.File.SaveBinaryDirect((ClientContext)ctx, webUrl.AbsolutePath + "/_cts/" + documentSetName + "/" + documentItem.File.Name, memStream, true);
                memStream.Close();
            }
        }

        //public static void AddMajorVersion(this DocumentSetTemplate documentSet, ListItem documentItem, string documentSetName, string folder, ContentTypeId docId)
        //{
        //    var ctx = (ClientContext)documentSet.Context;

        //    ctx.Load(documentItem.File.Versions);
        //    ctx.ExecuteQuery();

        //    string majorVersion = documentItem.File.MajorVersion.ToString() + ".0";

        //    ClientResult<Stream> documentStream = null;
        //    foreach (FileVersion fileVersion in documentItem.File.Versions)
        //    {
        //        if (fileVersion.VersionLabel.Equals(majorVersion))
        //        {
        //            documentStream = fileVersion.OpenBinaryStream();
        //            ctx.ExecuteQuery();
        //            break;
        //        }
        //    }

        //    string placeholderPage = "Placeholder";
        //    MemoryStream repo = new MemoryStream(Encoding.UTF8.GetBytes(placeholderPage));
        //    documentSet.DefaultDocuments.Add(folder + documentItem.File.Name, docId, repo.ToArray());
        //    documentSet.Update(true);
        //    ctx.ExecuteQuery();
        //    repo.Close();

        //    MemoryStream memStream = new MemoryStream();
        //    documentStream.Value.CopyTo(memStream);
        //    memStream.Seek(0, SeekOrigin.Begin);

        //    Uri url = new Uri(ctx.Url);

        //    Microsoft.SharePoint.Client.File.SaveBinaryDirect((ClientContext)ctx, url.AbsolutePath + "/_cts/" + documentSetName + "/" + documentItem.File.Name, memStream, true);
        //    memStream.Close();
        //}
    }
}
