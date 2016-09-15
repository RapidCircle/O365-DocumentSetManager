using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RapidCircle.SharePoint.DocumentSets
{
    public class DocumentSetManagerConfiguration
    {
        public string DocumentLibrary { get; set; }
        public Dictionary<string,string> FolderToDocumentSetMapping { get; set; }
        public List<string> ExcludedFolders { get; set; }
        public bool MajorVersionsOnly { get; set; }

        public DocumentSetManagerConfiguration(string documentLibrary, Dictionary<string, string> folderToDocumentSetMapping)
        {
            DocumentLibrary = documentLibrary;
            FolderToDocumentSetMapping = folderToDocumentSetMapping;
            MajorVersionsOnly = true;
            ExcludedFolders = new List<string>();
        }
    }
}
