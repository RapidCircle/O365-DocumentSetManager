using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using File = Microsoft.SharePoint.Client.File;

namespace RapidCircle.SharePoint.DocumentSets
{
    public class DocumentStream
    {
        private readonly ClientContext _ctx;
        private MemoryStream _latestVersion;
        private File _documentFile;

        public readonly string MajorVersionLabel;
        public readonly string FileName;

        public File Document { get; set; }

        public MemoryStream MajorVersion
        {
            get
            {
                if (_documentFile.UIVersionLabel.Equals(MajorVersionLabel)) //_documentFile.Versions.Count.Equals(0) && 
                    return LatestVersion;

                foreach (FileVersion fileVersion in _documentFile.Versions)
                {
                    if (fileVersion.VersionLabel.Equals(MajorVersionLabel))
                    {
                        var documentStream = fileVersion.OpenBinaryStream();
                        _ctx.ExecuteQuery();
                        Trace.TraceInformation($"- Major version {MajorVersionLabel} retrieved.");
                        return convertToMemoryStream(documentStream);
                    }
                }
                return null;
            }
        }

        public MemoryStream LatestVersion
        {
            get
            {
                var documentStream = _documentFile.OpenBinaryStream();
                _ctx.ExecuteQuery();

                return convertToMemoryStream(documentStream);
            }
        }

        public DocumentStream(ClientContext ctx, ListItem documentItem)
        {
            _ctx = ctx;

            _documentFile = documentItem.File;
            ctx.Load(_documentFile);
            ctx.Load(_documentFile.Versions);
            ctx.ExecuteQuery();

            MajorVersionLabel = _documentFile.MajorVersion + ".0";
            FileName = _documentFile.Name;
        }

        private MemoryStream convertToMemoryStream(ClientResult<Stream> documentStream)
        {
            MemoryStream documentMemoryStream = new MemoryStream();
            documentStream.Value.CopyTo(documentMemoryStream);
            documentMemoryStream.Position = 0;
            return documentMemoryStream;
        }
    }
}
