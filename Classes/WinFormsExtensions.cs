using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MobileNumbersDetailizationReportGenerator
{
    internal static class WinFormsExtensions
    {
        internal static void AppendLine(this TextBox source, string value = "\r\n")
        {
            if (source?.Text?.Length == 0)
                source.Text = value;
            else
                source.AppendText("\r\n" + value);
        }

        internal static void WriteAtFile(this string source, string filePath)
        {
            File.WriteAllText(
                filePath,
                source,
                Encoding.GetEncoding(1251));
        }

        internal static void WriteAtFile(this List<string> source, string filePath)
        {
            using (StreamWriter swExtLogFile = new StreamWriter(filePath, true))
            {
                foreach (var s in source.ToArray())
                {
                    swExtLogFile.WriteLine(s);
                }
            }
        }

        internal static void AppendAtFile(this string source, string filePath)
        {
            File.AppendAllText(
                filePath,
                source,
                Encoding.GetEncoding(1251));
        }

        //internal static void WriteAtFile(this List<string> listStrings, string filePath)
        //{
        //    File.WriteAllLines(
        //        filePath,
        //        listStrings,
        //        Encoding.GetEncoding(1251));
        //}
        internal static void AppendAtFile(this List<string> listStrings, string filePath)
        {
            File.AppendAllLines(
                filePath,
                listStrings,
                Encoding.GetEncoding(1251));
        }

        internal static string AsString(this List<string> source, string separator)
        {
            return string.Join(separator, source.ToArray());
        }

        internal static decimal TryParseAsInternetTrafic(this string source, string formResult)
        {

            string endSource = source?.Trim()?.ToUpper();
           
            if (endSource?.Length < 1|| !endSource.Contains(' '))
                return 0;

            string end = string.Empty;

            if (endSource.EndsWith("MB"))
            {
                end = "Mb";
            }
            else if (endSource.EndsWith("KB"))
            {
                end = "Kb";
            }
            else if (endSource.EndsWith("B"))
            {
                end = "b";
            }

            int.TryParse(endSource?.Remove(endSource.IndexOf(' ')), out int parsed);
            
            decimal result;
            switch (end)
            {
                case ("Mb"):
                    result = parsed * 1024 * 1024/MultiplierInternetTrafic.MultiplierInB(formResult);
                    break;
                case ("Kb"):
                    result = parsed * 1024 / MultiplierInternetTrafic.MultiplierInB(formResult);
                    break;
                case ("b"):
                    result = parsed / MultiplierInternetTrafic.MultiplierInB(formResult);
                    break;
                default:
                    result = 0;
                    break;
            }

            return Math.Round(result, 2);
        }
    }

}
