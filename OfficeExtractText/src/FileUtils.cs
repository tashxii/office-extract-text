using System.Collections.Generic;
using System.IO;
using System.Text;

namespace OfficeExtractText
{
    class FileUtils
    {
        public static List<string> MergeTextContents(string[] tempFiles)
        {
            var contents = new List<string>();
            foreach (var tempFile in tempFiles)
            {
                if (File.Exists(tempFile))
                {
                    contents.AddRange(File.ReadAllLines(tempFile, Encoding.Default));
                }
            }
            return contents;
        }

        public static void DeleteFiles(string[] tempFiles)
        {
            foreach (var tempFile in tempFiles)
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }
        }
    }
}
