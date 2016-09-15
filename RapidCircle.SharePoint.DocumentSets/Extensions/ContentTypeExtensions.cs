using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;

namespace RapidCircle.SharePoint.DocumentSets.Extensions
{
    public static class ContentTypeExtensions
    {
        public static ContentType GetByName(this ContentTypeCollection cts, string name)
        {
            var ctx = cts.Context;
            ctx.Load(cts);
            ctx.ExecuteQuery();
            return Enumerable.FirstOrDefault(cts, ct => ct.Name == name);
        }
    }
}
