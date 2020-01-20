using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
    /// <summary>
    /// Data From T-Factura's DB 
    /// </summary>
    public class MobileNumberOwner
    {

        public string MobileNumber { get; set; }
        public string FIO { get; set; }
        public string NAV { get; set; }
        public string Department { get; set; }
        public string ModelCompensation { get; set; }
        public string HoldingPeriod { get; set; }
    }
}
