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
    }
}
