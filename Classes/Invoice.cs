using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
    internal class Invoice
    {
        internal string invoiceFileName; // путь до текстового файла с детализацией
        internal string invoicePathToFile; // путь до текстового файла с детализацией
        internal string invoiceInternalHoldingNumber; //"Особовий рахунок"
        internal string invoiceNumber; //"Номер рахунку"
        internal string invoicePeriod; //"Розрахунковий період"

        internal double invoiceDeliveryCost; // Скидка навесь счет
        internal double invoiceDeliveryCostDiscount; // скидка на услугу детализ.счет в электронном виде
    }
}
