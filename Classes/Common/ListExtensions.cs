//using System.Collections.Generic;
//using System.IO;
//using System.Text;

using System.Collections.Generic;
using System.IO;
using System.Text;

namespace MobileNumbersDetailizationReportGenerator
{
    public static class ListExtensions
    {
        public static void WriteAtFile(this List<string> source, string filePath)
        {
            using (StreamWriter swExtLogFile = new StreamWriter(filePath, true))
            {
                foreach (var s in source.ToArray())
                {
                    swExtLogFile.WriteLine(s);
                }
            }
        }

        public static void AppendAtFile(this List<string> source, string filePath)
        {
            File.AppendAllLines(
                filePath,
                source,
                Encoding.GetEncoding(1251));
        }

        public static string AsString(this List<string> source, string separator)
        {
            return string.Join(separator, source.ToArray());
        }

        //public static void WriteAtFile(this List<string> listStrings, string filePath)
        //{
        //    File.WriteAllLines(
        //        filePath,
        //        listStrings,
        //        Encoding.GetEncoding(1251));
        //}
    }
}
