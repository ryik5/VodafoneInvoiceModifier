
namespace MobileNumbersDetailizationReportGenerator
{
    public struct Invoice
    {
        public string FileName { get; set; }//; // имя текстового файла с детализацией
        public string PathToFile { get; set; }//; // путь к текстовому файлу с детализацией
        public string InternalHoldingNumber { get; set; }//; //"Особовий рахунок"
        public string Nubmer { get; set; }//; //"Номер рахунку"
        public string Period { get; set; }//; //"Розрахунковий період"

        public double BillDeliveryValue { get; set; }//; // Стоимость Услуги доставки электронного счет
        public double DiscountOnBillDeliveryValue { get; set; }//; // скидка на услугу детализ.счет в электронном виде
    }
}
