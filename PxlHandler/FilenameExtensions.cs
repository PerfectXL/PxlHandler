using System.Text.RegularExpressions;

namespace PxlHandler {
    internal static class FilenameExtensions
    {
        public static string StripKnownExcelExtension(this string filename)
        {
            return Regex.Replace(filename, @"\.(?:xlm|xls|xlsb|xlsm|xlsx|xlt|xltx|xltm)$", "", RegexOptions.IgnoreCase);
        }
    }
}