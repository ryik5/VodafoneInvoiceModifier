namespace MobileNumbersDetailizationReportGenerator
{

    public class ParsedContractOfBill : StringOfDetalizationOfContractOfBill
    {
        public string contract { get; set; }
        public string numberOwner { get; set; }

        public string FIO { get; set; }
        public string NAV { get; set; }
        public string Department { get; set; }
    }
}
