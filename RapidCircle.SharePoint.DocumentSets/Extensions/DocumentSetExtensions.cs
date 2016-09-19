﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.DocumentSet;
using File = Microsoft.SharePoint.Client.File;

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
                memStream.Position = 0;
                //File.SaveBinaryDirect((ClientContext)ctx, webUrl.AbsolutePath + "/_cts/" + documentSetName + "/" + documentItem.File.Name, memStream, true);
                //memStream.Close();

                UploadFile((ClientContext)ctx, webUrl.AbsolutePath + "/_cts/" + documentSetName, memStream, documentItem.File.Name);
                //memStream.Close();
            }
        }

        private static void UploadFile(ClientContext ctx, string folderPath, MemoryStream memStream, string fileName)
        {
            Guid uploadId = Guid.NewGuid();
            File uploadFile;
            //int blockSize = 1000;
            int blockSize = 1 * 1024 * 1024;

            Folder uploadFolder = ctx.Web.GetFolderByServerRelativeUrl(folderPath);
            ctx.Load(uploadFolder, f => f.ServerRelativeUrl);
            ctx.ExecuteQuery();

            long fileSize = memStream.Length;

            if (fileSize <= blockSize)
            {
                // Use regular approach
                //using (FileStream fs = new FileStream(fileName, FileMode.Open))
                //{
                    FileCreationInformation fileInfo = new FileCreationInformation();
                    fileInfo.ContentStream = memStream;
                    fileInfo.Url = fileName;
                    fileInfo.Overwrite = true;
                    uploadFile = uploadFolder.Files.Add(fileInfo);
                    ctx.Load(uploadFile);
                    ctx.ExecuteQuery();
                    // return the file object for the uploaded file
                    //return uploadFile;
                //}
            }
            else
            {

                ClientResult<long> bytesUploaded = null;
                FileStream fs = null;
                try
                {
                    //fs = System.IO.File.Open(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using (BinaryReader br = new BinaryReader(memStream))
                    {
                        byte[] buffer = new byte[blockSize];
                        Byte[] lastBuffer = null;
                        long fileoffset = 0;
                        long totalBytesRead = 0;
                        int bytesRead;
                        bool first = true;
                        bool last = false;

                        // Read data from filesystem in blocks 
                        while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                        {
                            totalBytesRead = totalBytesRead + bytesRead;

                            // We've reached the end of the file
                            if (totalBytesRead == fileSize)
                            {
                                last = true;
                                // Copy to a new buffer that has the correct size
                                lastBuffer = new byte[bytesRead];
                                Array.Copy(buffer, 0, lastBuffer, 0, bytesRead);
                            }

                            if (first)
                            {
                                using (MemoryStream contentStream = new MemoryStream())
                                {
                                    // Add an empty file.
                                    FileCreationInformation fileInfo = new FileCreationInformation();
                                    fileInfo.ContentStream = contentStream;
                                    fileInfo.Url = fileName;
                                    fileInfo.Overwrite = true;
                                    uploadFile = uploadFolder.Files.Add(fileInfo);

                                    // Start upload by uploading the first slice. 
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Call the start upload method on the first slice
                                        bytesUploaded = uploadFile.StartUpload(uploadId, s);
                                        ctx.ExecuteQuery();
                                        // fileoffset is the pointer where the next slice will be added
                                        fileoffset = bytesUploaded.Value;
                                    }

                                    // we can only start the upload once
                                    first = false;
                                }
                            }
                            else
                            {
                                // Get a reference to our file
                                uploadFile =
                                    ctx.Web.GetFileByServerRelativeUrl(uploadFolder.ServerRelativeUrl +
                                                                       System.IO.Path.AltDirectorySeparatorChar +
                                                                       fileName);

                                if (last)
                                {
                                    // Is this the last slice of data?
                                    using (MemoryStream s = new MemoryStream(lastBuffer))
                                    {
                                        // End sliced upload by calling FinishUpload
                                        uploadFile = uploadFile.FinishUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();

                                        // return the file object for the uploaded file
                                        //return uploadFile;
                                    }
                                }
                                else
                                {
                                    using (MemoryStream s = new MemoryStream(buffer))
                                    {
                                        // Continue sliced upload
                                        bytesUploaded = uploadFile.ContinueUpload(uploadId, fileoffset, s);
                                        ctx.ExecuteQuery();
                                        // update fileoffset for the next slice
                                        fileoffset = bytesUploaded.Value;
                                    }
                                }
                            }

                        } // while ((bytesRead = br.Read(buffer, 0, buffer.Length)) > 0)
                    }
                }
                finally
                {
                    if (fs != null)
                    {
                        fs.Dispose();
                    }
                }

            }

        }
    }
}
