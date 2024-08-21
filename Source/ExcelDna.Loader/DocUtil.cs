using System;
using System.IO;
using System.Linq;

namespace ExcelDna.Loader
{
    internal static class DocUtil
    {
        static string tempDocDir;

        public static string FixHelpTopic(string helpTopic)
        {
            string result = helpTopic;

            // Make HelpTopic without full path relative to xllPath
            if (string.IsNullOrEmpty(helpTopic))
            {
                return result;
            }

            // If http url does not end with !0 it is appended.
            // I don't think https is supported, but it should not be considered an 'unrooted' path anyway.
            // I could not get file:/// working (only checked with Excel 2013)
            if (helpTopic.StartsWith("http://") || helpTopic.StartsWith("https://") || helpTopic.StartsWith("file://"))
            {
                if (!helpTopic.EndsWith("!0"))
                {
                    result = helpTopic + "!0";
                }
            }
            else if (!Path.IsPathRooted(helpTopic))
            {
                result = TryFixUnpacked(helpTopic);
                if (!Path.IsPathRooted(result))
                    result = TryFixPacked(result);
            }

            return result;
        }

        public static void Clear()
        {
            try
            {
                if (tempDocDir != null)
                    Directory.Delete(tempDocDir, true);
            }
            catch
            {
            }
        }

        private static string TryFixUnpacked(string helpTopic)
        {
            string dir = Path.GetDirectoryName(XlAddIn.PathXll);
            if (File.Exists(Path.Combine(dir, GetFileName(helpTopic))))
                return Path.Combine(dir, helpTopic);

            return helpTopic;
        }

        private static string TryFixPacked(string helpTopic)
        {
            if (tempDocDir == null)
                tempDocDir = Path.Combine(XlAddIn.TempDirPath, Guid.NewGuid().ToString());

            Directory.CreateDirectory(tempDocDir);

            string fileName = GetFileName(helpTopic);
            string filePath = Path.Combine(tempDocDir, fileName);

            if (!File.Exists(filePath))
            {
                byte[] data = XlAddIn.GetResourceBytes(fileName, 6);
                if (data != null)
                    File.WriteAllBytes(filePath, data);
            }

            if (File.Exists(filePath))
                return Path.Combine(tempDocDir, helpTopic);

            return helpTopic;
        }

        private static string GetFileName(string helpTopic)
        {
            return helpTopic.Split('!').First();
        }
    }
}
