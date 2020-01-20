using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
    public interface IDetalizationParseable<T>
    {
        void Parse();

     //   List<T> InputList { get; set; }
    }
}
