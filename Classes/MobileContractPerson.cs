using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
    public class MobileContractPerson
    {
        internal string ownerName = "";
        internal string contractName = "";
        internal string mobNumberName = "";
        internal string tarifPackageName = "";
        internal double monthCost = 0;
        internal double roming = 0;
        internal double discount = 0;
        internal double totalCost = 0;
        internal double tax = 0;
        internal double pF = 0;
        internal double totalCostWithTax = 0;
        internal double totalCostWithoutTaxBeforDiscount = 0;
        internal double romingData = 0;
        internal double extraServiceOrdered = 0;
        internal double extraInternetOrdered = 0;
        internal double outToCity = 0;
        internal double extraService = 0;
        internal double content = 0;
        internal string dateBillStart = "";
        internal string dateBillEnd = "";

        internal string NAV = "";
        internal string orgUnit = "";
        internal string startDate;
        internal string modelCompensation = "";
        internal double payOwner = 0;
        internal bool isUsed = false;
        internal bool isUnblocked = false;
    }
}
