using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{

    public class ParsedBill //: IParseable
    {
        public List<ServiceOfBill> ServicesOfHeaderOfBill { get; set; }

        public List<ContractOfBill> ContractsOfBill { get; set; }

        public ParsedBill() { }

        //    public void Parse()
        //  {        }

    }


    /// <summary>
    /// only fully parsed contract with header and body detalization
    /// </summary>
    public class ContractOfBill
    {
        public HeaderOfContractOfBill Header { get; private set; }

        public ServicesOfContractOfBill ServicesOfContract { get; private set; }

        public DetalizationOfContractOfBill DetalizationOfContract { get; private set; }

        public List<string> Source { get; private set; }

        public ContractOfBill(List<string> source) { Source = source; }

        public ContractOfBill(HeaderOfContractOfBill header, ServicesOfContractOfBill services, DetalizationOfContractOfBill detalization)
        {
            Header = header;
            ServicesOfContract = services;
            DetalizationOfContract = detalization;
        }
    }



    public class ServicesOfContractOfBill : AbstractPartOfContractDetalization<ServiceOfBill>//, IParseable
    {
        public ServicesOfContractOfBill(List<string> source) : base(source) { }

        public ServicesOfContractOfBill(List<ServiceOfBill> list )
        { Output = list; }
       
        //public override void Parse()
        //{
        //    if (!(Output?.Count > 0))
        //    {
        //        Output = new List<ServiceOfBill>();
        //    }

        //    //todo - parse header of Contract
        //}
    }


    public class DetalizationOfContractOfBill : AbstractPartOfContractDetalization<StringOfDetalizationOfContractOfBill>//, IParseable
    {

        public DetalizationOfContractOfBill(List<string> source) : base(source) { }

        public DetalizationOfContractOfBill() { }
        //public override void Parse()
        //{
        //    if (!(Output?.Count > 0))
        //    {
        //        Output = new List<StringOfDetalizationOfContractOfBill>();
        //    }

        //    //todo - parse Body

        //}
    }




    public class HeaderOfContractOfBill
    {
        public HeaderOfContractOfBill() { }

        public HeaderOfContractOfBill(string id, string number, string tarif)
        {
            ContractId = id;
            MobileNumber = number;
            TarifPackage = tarif;
        }

        public HeaderOfContractOfBill(List<string> source)
        {
            Source = source;
        }

        public List<string> Source { get; private set; }

        public string ContractId { get; private set; }

        public string MobileNumber { get; private set; }

        public string TarifPackage { get; private set; }

    }

    public class StringOfDetalizationOfContractOfBill
    {
        public string ServiceName { get; set; }
        public string NumberTarget { get; set; }
        public string Date { get; set; }
        public string Time { get; set; }
        public string DurationA { get; set; }
        public string DurationB { get; set; }
        public string Cost { get; set; }
    }

    public class ServiceOfBill
    {
        public string Name { get; set; }

        public double Amount { get; set; }

        public ServiceOfBill(string name, double amount)
        {
            Name = name;
            Amount = amount;
        }
    }

}
