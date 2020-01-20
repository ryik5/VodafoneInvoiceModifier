using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
    //todo
    //Remove it
    public class ParsedContractOfBill : ParsedStringOfBodyOfContract
    {
        public string contract { get; set; }
        public string numberOwner { get; set; }


        public Contract Contract { get; set; }
        public string FIO { get; set; }
        public string NAV { get; set; }
        public string Department { get; set; }
    }
}
