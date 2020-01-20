using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{

    /// <summary>
    /// Contract,List<string> DataIn, Parse()
    /// </summary>
    abstract public class AbstractParsedPartOfContractDetalization: IParseable
    {
        public Contract Contract { get; set; }

        public List<string> DataIn { get; set; }

        public virtual void Parse() { }

        public AbstractParsedPartOfContractDetalization(List<string> source)
        {
            this.DataIn = source;
        }
    }

}
