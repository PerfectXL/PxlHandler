using System;
using System.Linq;
using System.Web;

namespace PxlHandler
{
    internal static class UriExtensions
    {
        public static PxlPath GetPathParts(this Uri uri)
        {
            var parts = uri.AbsolutePath.Split(new[] {'/'}, StringSplitOptions.RemoveEmptyEntries).Select(HttpUtility.UrlDecode).ToArray();
            return parts.Length < 2 ? PxlPath.Invalid() : new PxlPath(parts[0], parts[1], parts.Skip(2).FirstOrDefault());
        }
    }
}