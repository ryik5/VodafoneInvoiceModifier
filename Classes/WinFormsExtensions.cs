using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MobileNumbersDetailizationReportGenerator
{
    public static class WinFormsExtensions
    {
        public static void AppendLine(this TextBox source, string value = "\r\n")
        {
            if (source?.Text?.Length == 0)
                source.Text = value;
            else
                source.AppendText($"{Environment.NewLine} {value}");
        }
        

        /// <summary>
        /// waiting string as '200 Mb'
        /// </summary>
        /// <param name="source"></param>
        /// <param name="formResult"></param>
        /// <returns></returns>
        public static decimal ToInternetTrafic(this string source, string formResult)
        {

            string text = source?.Trim()?.ToUpper();

            if (text?.Length < 1 || !text.Contains(' '))
                return 0;

            string end = string.Empty;

            if (text.EndsWith("MB"))
            {
                end = "Mb";
            }
            else if (text.EndsWith("KB"))
            {
                end = "Kb";
            }
            else if (text.EndsWith("B"))
            {
                end = "b";
            }

            int.TryParse(text?.Remove(text.IndexOf(' ')), out int parsed);

            decimal result;
            switch (end)
            {
                case ("Mb"):
                    result = parsed * 1024 * 1024 / MultiplierInternetTrafic.MultiplierInB(formResult);
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
