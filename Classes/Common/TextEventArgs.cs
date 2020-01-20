using System;

namespace MobileNumbersDetailizationReportGenerator
{
    public class TextEventArgs : EventArgs
    {
        // Сообщение
        public string Message { get; private set; }

        public TextEventArgs(string message)
        {
            Message = message;
        }
    }
}
