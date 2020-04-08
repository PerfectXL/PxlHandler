using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.Win32;

namespace PxlHandler
{
    internal class Program
    {
        public static void Main(string[] args)
        {
            RegisterUriScheme();

            if (!args.Any())
            {
                return;
            }

            try
            {
                PxlResult result = FollowLink(args[0]);
                PrintResult(result);

                if (result == PxlResult.Success)
                {
                    ProcessHelper.BringProcessToFront("EXCEL");
                }
            }
            catch (Exception exception)
            {
                Console.WriteLine($"Error: {exception.Message} Line: {exception.Source}");
            }

            Console.ReadKey();
        }

        public static void RegisterUriScheme()
        {
            var applicationLocation = GetAssemblyFullPath();
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey("SOFTWARE\\Classes\\pxl"))
            {
                Debug.Assert(key != null, nameof(key) + " != null");
                key.SetValue("", "PerfectXL Handler");
                key.SetValue("URL Protocol", "");

                using (RegistryKey defaultIcon = key.CreateSubKey("DefaultIcon"))
                {
                    Debug.Assert(defaultIcon != null, nameof(defaultIcon) + " != null");
                    defaultIcon.SetValue("", $"{applicationLocation},1");
                }

                using (RegistryKey commandKey = key.CreateSubKey(@"shell\open\command"))
                {
                    Debug.Assert(commandKey != null, nameof(commandKey) + " != null");
                    commandKey.SetValue("", $"\"{applicationLocation}\" \"%1\"");
                }
            }
        }

        private static PxlResult FollowLink(string uriString)
        {
            if (!Uri.TryCreate(uriString, UriKind.Absolute, out Uri uri))
            {
                return PxlResult.InvalidUri;
            }

            if (uri.Scheme != "pxl")
            {
                return PxlResult.InvalidScheme;
            }

            if (uri.Host != "jump-to-cell")
            {
                return PxlResult.InvalidHost;
            }

            PxlPath pxlPath = uri.GetPathParts();
            Console.WriteLine(pxlPath);
            if (!pxlPath.IsValid)
            {
                return PxlResult.InvalidPath;
            }

            return PxlProtocolHelper.JumpToCell(pxlPath) ? PxlResult.Success : PxlResult.JumpToCellFailed;
        }

        private static string GetAssemblyFullPath()
        {
            var codeBase = Assembly.GetExecutingAssembly().CodeBase;
            var uri = new UriBuilder(codeBase);
            var path = Uri.UnescapeDataString(uri.Path);
            return Path.GetFullPath(path);
        }

        private static void PrintResult(PxlResult result)
        {
            switch (result)
            {
                case PxlResult.Unknown:
                    Console.WriteLine("Result is unknown.");
                    break;
                case PxlResult.Success:
                    Console.WriteLine("Success!");
                    break;
                case PxlResult.InvalidUri:
                    Console.WriteLine("Invalid URI.");
                    break;
                case PxlResult.InvalidScheme:
                    Console.WriteLine("We handle only the PXL scheme.");
                    break;
                case PxlResult.InvalidHost:
                    Console.WriteLine("Currently only pxl://jump-to-cell is implemented.");
                    break;
                case PxlResult.InvalidPath:
                    Console.WriteLine("Invalid path. Provide the file name, the worksheet name and optionally the range.");
                    break;
                case PxlResult.JumpToCellFailed:
                    Console.WriteLine("Could not jump to the given location.");
                    break;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
    }
}