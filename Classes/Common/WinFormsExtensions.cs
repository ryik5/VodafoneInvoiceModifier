using System;
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
        /// waiting source as '20 Gb' or '20 Mb' or '20 Kb' or '20 b' 
        /// formResult as 'Gb' or 'Mb' or 'Kb' or 'b'
        /// </summary>
        /// <param name="source">as '20 Gb' or '20 Mb' or '20 Kb' or '20 b'</param>
        /// <param name="formResult">as 'Gb' or 'Mb' or 'Kb' or 'b'</param>
        /// <returns>20</returns>
        public static decimal ToInternetTrafic(this string source, string formResult)
        {
            string text = source?.Trim()?.ToUpper();

            if (!(text?.Length > 0))
            { return 0; }

            if (MultiplierInternetTrafic.Multiplier(formResult) < 1)
            {
                throw new Exception("Wrong multiplier!");
            }

            string end;

            if (text.EndsWith("GB"))
            {
                end = "GB";
            }
            else if (text.EndsWith("MB"))
            {
                end = "MB";
            }
            else if (text.EndsWith("KB"))
            {
                end = "KB";
            }
            else if (text.EndsWith("B"))
            {
                end = "B";
            }
            else
            {
                return 0;
            }

            decimal parsed = 0;
            bool inputCorrect = decimal.TryParse(text.Replace(end, "").Trim(), out parsed);

            if (!inputCorrect)
            { return 0; }

            decimal result;

            switch (end)
            {
                case ("GB"):
                    result = parsed * 1024 * 1024 * 1024 / MultiplierInternetTrafic.Multiplier(formResult);
                    break;
                case ("MB"):
                    result = parsed * 1024 * 1024 / MultiplierInternetTrafic.Multiplier(formResult);
                    break;
                case ("KB"):
                    result = parsed * 1024 / MultiplierInternetTrafic.Multiplier(formResult);
                    break;
                case ("B"):
                    result = parsed / MultiplierInternetTrafic.Multiplier(formResult);
                    break;
                default:
                    result = 0;
                    break;
            }

            return Math.Round(result, 3);
        }
    }
}
