
namespace MobileNumbersDetailizationReportGenerator
{
    public struct Invoice
    {
        public string invoiceFileName { get; set; }//; // путь до текстового файла с детализацией
        public string invoicePathToFile { get; set; }//; // путь до текстового файла с детализацией
        public string invoiceInternalHoldingNumber { get; set; }//; //"Особовий рахунок"
        public string invoiceNumber { get; set; }//; //"Номер рахунку"
        public string invoicePeriod { get; set; }//; //"Розрахунковий період"

        public double invoiceDeliveryCost { get; set; }//; // Скидка навесь счет
        public double invoiceDeliveryCostDiscount { get; set; }//; // скидка на услугу детализ.счет в электронном виде
    }
}
