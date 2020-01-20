using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
    abstract public class ParsedPartOfContractDetalization
    {
        public Contract Contract { get; set; }

        public List<string> DataIn { get; set; }

        public virtual void Parse() { }

        public ParsedPartOfContractDetalization(List<string> source)
        {
            this.DataIn = source;
        }
    }

    public class ParsedHeaderOfContract : ParsedPartOfContractDetalization, IDetalizationParseable<string>
    {
        public List<ServiceInDetalization> ServicesInHeader { get; set; }

        public ParsedHeaderOfContract(List<string> source) : base(source) { }

        public override void Parse()
        {

        }
    }

    public class ParsedBodyOfContract : ParsedPartOfContractDetalization, IDetalizationParseable<string>
    {

        public List<ParsedStringOfBodyOfContract> ParsedBody { get; set; }

        public ParsedBodyOfContract(List<string> source) : base(source) { }

        public override void Parse()
        {
            //todo - parse Body
        }
    }

    public class ParsedStringOfBodyOfContract
    {
        public string ServiceName { get; set; }
        public string NumberTarget { get; set; }
        public string Date { get; set; }
        public string Time { get; set; }
        public string DurationA { get; set; }
        public string DurationB { get; set; }
        public string Cost { get; set; }

        public string DataIn{ get; set; }
    }

    public class ParsedContractOfBillDetalization
    {
        public Contract Contract { get; set; }
        public ParsedBodyOfContract ParsedBodyOfContract { get; set; }
        public ParsedHeaderOfContract ParsedHeaderOfContract { get; set; }
    }

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
