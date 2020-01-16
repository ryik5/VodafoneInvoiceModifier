
namespace MobileNumbersDetailizationReportGenerator
{
    public struct MobileContractPerson
    {
        public string OwnerName { get; set; }// = "";
        public string Сontract { get; set; }//= "";
        public string CellNumber { get; set; }//= "";
        public string NumberTarifPackageName { get; set; }// = "";
        public double NumberMonthCost { get; set; }//= 0;
        public double RoamingSummary { get; set; }//= 0;
        public double ContractDiscount { get; set; }// = 0;
        public double TotalCost { get; set; }// = 0;//Zatraty na Contract do nalogov
        public double TaxCost { get; set; }//= 0;//nalog na dob.stoimost
        public double PFCost { get; set; }//= 0;//Pensioniy fond
        public double TotalCostWithTax { get; set; }//= 0;
        public double RomingDetalization { get; set; }//= 0;
        public double PaidExtraOfTarifPackageServices { get; set; }//= 0;
        public double PaidExtraOfTarifPackageInternetService { get; set; }//= 0;
        public double PaidExtraOfTarifOutCallsToCity { get; set; }// = 0;
        public double ExtraService { get; set; }//= 0;
        public double PaidExtraContentOfTarifPackage { get; set; }//= 0;
        public string StartDatePeriodBill { get; set; }// = "";
        public string EndDayPeriodBill { get; set; }//= "";

        public string NAV { get; set; }// = "";
        public string Department { get; set; }//= "";
        public string StartDayOfModelCompensation { get; set; }//;
        public string ModelCompensation { get; set; }//= "";
        public double payOwner { get; set; }//= 0;
        public bool isUsedNumber { get; set; }//= false;
        public bool isUnblockedNumber { get; set; }//= false;
    }
}
