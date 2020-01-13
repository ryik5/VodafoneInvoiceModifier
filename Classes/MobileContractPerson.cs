using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
    public struct MobileContractPerson
    {
        public string ownerName { get; set; }// = "";
        public string contractName { get; set; }//= "";
        public string mobNumberName { get; set; }//= "";
        public string tarifPackageName { get; set; }// = "";
        public double monthCost { get; set; }//= 0;
        public double roming { get; set; }//= 0;
        public double discount { get; set; }// = 0;
        public double totalCost { get; set; }// = 0;
        public double tax { get; set; }//= 0;
        public double pF { get; set; }//= 0;
        public double totalCostWithTax { get; set; }//= 0;
        public double totalCostWithoutTaxBeforDiscount { get; set; }// = 0;
        public double romingData { get; set; }//= 0;
        public double extraServiceOrdered { get; set; }//= 0;
        public double extraInternetOrdered { get; set; }//= 0;
        public double outToCity { get; set; }// = 0;
        public double extraService { get; set; }//= 0;
        public double content { get; set; }//= 0;
        public string dateBillStart { get; set; }// = "";
        public string dateBillEnd { get; set; }//= "";

        public string NAV { get; set; }// = "";
        public string orgUnit { get; set; }//= "";
        public string startDate { get; set; }//;
        public string modelCompensation { get; set; }//= "";
        public double payOwner { get; set; }//= 0;
        public bool isUsed { get; set; }//= false;
        public bool isUnblocked { get; set; }//= false;
    }
}
