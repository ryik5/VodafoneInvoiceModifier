using System.Collections.Generic;

namespace MobileNumbersDetailizationReportGenerator
{

    /// <summary>
    /// Contract,List<string> DataIn, Parse()
    /// </summary>
    abstract public class AbstractPartOfContractDetalization<T>//: IParseable
    {
        public HeaderOfContractOfBill Contract { get; set; }

        public List<string> Source { get; set; }

        public List<T> Output { get; protected set; }

        public AbstractPartOfContractDetalization() { }

        public AbstractPartOfContractDetalization(List<string> source)
        {
            Source = source;
        }
    }
}
