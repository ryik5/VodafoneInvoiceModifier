using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MobileNumbersDetailizationReportGenerator
{
   public static class StringExtensions
    {
        public static void AppendAtFile(this string source, string filePath)
        {
            File.AppendAllText(
                filePath,
                source,
                Encoding.GetEncoding(1251));
        }

        public static void WriteAtFile(this string source, string filePath)
        {
            File.WriteAllText(
                filePath,
                source,
                Encoding.GetEncoding(1251));
        }

        public static string[] ExpandArray(this string[] array, string addToList)
        {
            if (!(addToList?.Length > 0) || !(array?.Length > 0))
            { return array; }

            List<string> list = array.ToList();
            list.Add(addToList);
            string[] temp = list.ToArray();
           
            return temp;
        }

    }
}
