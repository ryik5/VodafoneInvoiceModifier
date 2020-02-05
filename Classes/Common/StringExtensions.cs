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

        /// <summary>
        /// Expand array upto an extra cell
        /// </summary>
        /// <param name="source">array which need to expand</param>
        /// <param name="extraCell">an extra cell</param>
        /// <returns></returns>
        public static string[] ExpandArray(this string[] source, string extraCell)
        {
            if (!(extraCell?.Length > 0) || !(source?.Length > 0))
            { return source; }

            List<string> list = source.ToList();
            list.Add(extraCell);
            string[] temp = list.ToArray();
           
            return temp;
        }

        /// <summary>
        /// Expand array upto an extra cell
        /// </summary>
        /// <param name="source">array which need to expand</param>
        /// <param name="extraCell">an extra cell</param>
        /// <returns></returns>
        public static int[] ExpandArray(this int[] source, int extraCell)
        {
            if (!(source?.Length > 0))
            { return source; }

            List<int> list = source.ToList();
            list.Add(extraCell);
            int[] temp = list.ToArray();

            return temp;
        }

    }
}
