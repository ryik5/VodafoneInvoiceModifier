using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

    }
}
