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

        //Access to Control from other threads
        public static string OpenFileDialogReturnPath(this OpenFileDialog ofd) //Return its name 
        {
            ofd.FileName = @"";
            ofd.Filter = Properties.Resources.OpenDialogTextFiles;
            ofd.ShowDialog();
            string filePath = ofd.FileName;

            return filePath;
        }

    }
}
