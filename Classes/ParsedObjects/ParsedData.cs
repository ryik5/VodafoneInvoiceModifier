using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
 
    public class ParsedBill : IParseable
    {
        public List<ParsedContract> ParsedContracts { get; set; }

        public List<ParsedService> ParsedServicesOfBill { get; set; }

        public List<string> Source;

        public ParsedBill(List<string> source)
        { this.Source = source; }

        public void Parse()
        {

        }

    }


    /// <summary>
    /// only fully parsed contract with header and body detalization
    /// </summary>
    public class ParsedContract
    {
        public Contract Contract { get; set; }

        public ParsedHeaderOfContract ParsedHeaderOfContract { get; set; }

        public ParsedBodyOfContract ParsedBodyOfContract { get; set; }

    }



    public class ParsedHeaderOfContract : AbstractParsedPartOfContractDetalization, IParseable
    {
        public List<ParsedService> ParsedServicesOfContract { get; set; }

        public ParsedHeaderOfContract(List<string> source) : base(source) { }

        public override void Parse()
        {

        }
    }

    public class ParsedBodyOfContract : AbstractParsedPartOfContractDetalization, IParseable
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

        public string RawData { get; set; }
    }


    public class Contract
    {
        public string MobileNumber { get; set; }

        public string ContractId { get; set; }

        public string TarifPackage { get; set; }

    }

    public class ParsedService
    {
        public string Name { get; set; }

        public double Amount { get; set; }
    }


}
