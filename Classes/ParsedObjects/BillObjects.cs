using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;

namespace MobileNumbersDetailizationReportGenerator
{

    public class ParsedBill 
    {
        public ServicesOfBill ServicesOfHeaderOfBill { get; set; }

        public ParsedBill() { }
    }

    /// <summary>
    /// only fully parsed contract with header and body detalization
    /// </summary>
    public class ContractOfBill
    {
        public HeaderOfContractOfBill Header { get; set; }

        public ServicesOfBill ServicesOfContract { get; set; }

        public DetalizationOfContractOfBill DetalizationOfContract { get; set; }

        public ContractOfBill(HeaderOfContractOfBill header, ServicesOfBill services, DetalizationOfContractOfBill detalization)
        {
            Header = header;
            ServicesOfContract = services;
            DetalizationOfContract = detalization;
        }

        public ContractOfBill(ContractOfBill contract)
        {
            Header = contract?.Header;
            ServicesOfContract = contract?.ServicesOfContract;
            DetalizationOfContract = contract?.DetalizationOfContract;
        }
    }

    public class ContractsRawOfBill : IEnumerable<List<string>>
    {
        private ContractRawListNode first;

        public void Add(List<string> list)
        {
            if (first == null)
                first = new ContractRawListNode
                {
                    Value = list
                };
            else
            {
                var node = this.first;
                while (node.Next != null)
                    node = node.Next;

                node.Next = new ContractRawListNode
                {
                    Value = list
                };
            }
        }

        public IEnumerator<List<string>> GetEnumerator()
        {
            for (var node = first; node != null; node = node.Next)
            {
                yield return node.Value;
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        private class ContractRawListNode
        {
            public List<string> Value { get; set; }
            public ContractRawListNode Next { get; set; }
        }
    }

    public class ServicesOfBill : AbstractPartOfContractDetalization<ServiceOfBill>//, IParseable
    {

        public ServicesOfBill(List<ServiceOfBill> list)
        { Output = list; }

        public ServicesOfBill(ServicesOfBill services)
        { Output = services.Output; }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var s in Output)
            {
                sb.AppendLine($"{s.Name.PadRight(90)}\t{s.IsMain.ToString().PadRight(8)}\t{s.Amount.ToString().PadRight(15)}");
            }

            return $"{sb.ToString()}\n";
        }
    }


    public class DetalizationOfContractOfBill : AbstractPartOfContractDetalization<StringOfDetalizationOfContractOfBill>//, IParseable
    {
        public DetalizationOfContractOfBill() { }
        public DetalizationOfContractOfBill(DetalizationOfContractOfBill detalization)
        { Output = detalization.Output; }

        public DetalizationOfContractOfBill(List<string> source) : base(source) { }

        public DetalizationOfContractOfBill(List<StringOfDetalizationOfContractOfBill> list)
        {
            Output = list;
        }

        public override string ToString()
        {
            if (Output == null)
                return null;

            StringBuilder sb = new StringBuilder();
            foreach (var s in Output)
            {
                sb.AppendLine(
                    $"{s.ServiceName.PadRight(80)}\t{s.NumberTarget.PadRight(14)}\t{s.Date.PadRight(10)}\t" +
                    $"{s.Time.PadRight(8)}\t{s.DurationA.PadRight(8)}\t{s.DurationB.PadRight(8)}\t{s.Cost.PadRight(10)}"
                    );
            }

            return $"{sb.ToString()}\n";
        }
    }


    public class HeaderOfContractOfBill
    {
        public HeaderOfContractOfBill() { }

        public HeaderOfContractOfBill(HeaderOfContractOfBill header)
        {
            ContractId = header.ContractId;
            MobileNumber = header.MobileNumber;
            TarifPackage = header.TarifPackage;
        }

        public HeaderOfContractOfBill(string id, string number, string tarif)
        {
            ContractId = id;
            MobileNumber = number;
            TarifPackage = tarif;
        }

        public string ContractId { get; private set; }

        public string MobileNumber { get; private set; }

        public string TarifPackage { get; private set; }

        public override string ToString()
        {
            return $"{ContractId.PadRight(20)}\t{MobileNumber.PadRight(20)}\t{TarifPackage.PadRight(40)}"; ;
        }
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

        public override string ToString()
        {
            return $"{ServiceName.PadRight(80)}\t{NumberTarget.PadRight(14)}\t{Date.PadRight(10)}\t" +
                $"{Time.PadRight(8)}\t{DurationA.PadRight(8)}\t{DurationB.PadRight(8)}\t{Cost.PadRight(10)}";
        }
    }


    public class ServiceOfBill
    {
        public string Name { get; set; }

        public double Amount { get; set; }

        public bool IsMain { get; set; }

        public ServiceOfBill(string name, double amount, bool isMain = false)
        {
            Name = name;
            Amount = amount;
            IsMain = isMain;
        }

        public override string ToString()
        {
            return $"{Name}\t{IsMain}\t{Amount}";
        }
    }
}
