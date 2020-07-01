using System;
using System.IO;
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
        public static string OpenFileDialogReturnPath(this OpenFileDialog ofd, string title) //Return its name 
        {
            if (ofd == null)
            {
                ofd = new OpenFileDialog
                {
                    Title = title,
                    FileName = @"",
                    Filter = Properties.Resources.OpenDialogTextFiles
                };
            }
                    
            ofd.ShowDialog();

            string filePath = ofd.FileName;

            return filePath;
        }

        static readonly object obj = new object();
        public static void Logger(LogTypes typo, string Event)
        {
            string path = System.Reflection.Assembly.GetExecutingAssembly().Location;
            string pathToLogDir, pathToLog;
            try
            {
                pathToLogDir = Path.Combine(Path.GetDirectoryName(path), $"logs");
                if (!Directory.Exists(pathToLogDir))
                    Directory.CreateDirectory(pathToLogDir);

                pathToLog = Path.Combine(pathToLogDir, $"{DateTime.Now.ToString("yyyy-MM-dd")}.log");
                lock (obj)
                {
                    using (StreamWriter writer = new StreamWriter(pathToLog, true))
                    {

                        writer.WriteLine($"{DateTime.Now.ToString("yyyy.MM.dd|hh:mm:ss")}|{typo}|{Event}");
                        writer.Flush();
                    }
                }
            }
            catch (Exception err)
            {
                pathToLog = Path.Combine(Path.GetDirectoryName(path), Path.GetFileNameWithoutExtension(path) + ".log");
                lock (obj)
                {
                    using (StreamWriter writer = new StreamWriter(pathToLog, true))
                    {
                        writer.WriteLine($"{DateTime.Now.ToString("yyyy.MM.dd|hh:mm:ss")}|{err.ToString()}");
                        writer.Flush();
                    }
                }
            }
        }

    }
        /// <summary>
        /// Info, Trace, Debug, Warn, Error, Fatal
        /// </summary>
        public enum LogTypes
        {
            Info = 0,
            Trace = 2,
            Debug = 4,
            Warn = 8,
            Error = 16,
            Fatal = 32
        }}
