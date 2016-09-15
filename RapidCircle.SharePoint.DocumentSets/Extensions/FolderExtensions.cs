using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace RapidCircle.SharePoint.DocumentSets.Extensions
{
    public static class FolderExtensions
    {
        public static Folder GetByName(this FolderCollection folders, string name)
        {
            var ctx = folders.Context;
            ctx.Load(folders);
            ctx.ExecuteQuery();
            return Enumerable.FirstOrDefault(folders, fldr => fldr.Name == name);
        }
    }
}
