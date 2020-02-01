using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{

    public class ParsedBill //: IParseable
    {
        List<string> wholeBill { get; set; }
        public List<ServiceOfBill> ServicesOfHeaderOfBill { get; set; }

        public List<ContractOfBill> ContractsOfBill { get; set; }

        public ParsedBill() { }

        public ParsedBill(List<string> wholeBill)
        {
            this.wholeBill = wholeBill;
        }
    }


    /// <summary>
    /// only fully parsed contract with header and body detalization
    /// </summary>
    public class ContractOfBill
    {
        public HeaderOfContractOfBill Header { get;  set; }

        public ServicesOfBill ServicesOfContract { get;  set; }

        public DetalizationOfContractOfBill DetalizationOfContract { get;  set; }

        public List<string> Source { get; private set; }


        public ContractOfBill() {
            Header = new HeaderOfContractOfBill();
            ServicesOfContract = new ServicesOfBill();
            DetalizationOfContract = new DetalizationOfContractOfBill();
        }

        public ContractOfBill(List<string> source) { Source = source; }
       
        public ContractOfBill(HeaderOfContractOfBill header, ServicesOfBill services, DetalizationOfContractOfBill detalization)
        {
            Header = header;
            ServicesOfContract = services;
            DetalizationOfContract = detalization;
        }
        
        public ContractOfBill(ContractOfBill contract)
        {
            Header = contract.Header;
            ServicesOfContract = contract.ServicesOfContract;
            DetalizationOfContract = contract.DetalizationOfContract;
        }
    }

    //public class ContractsRawOfBill
    //{
    //    static object check;

    //    public ContractsRawOfBill()
    //    {
    //        this.Contracts = new List<ContractRawList>();
    //    }

    //    public ContractsRawOfBill(ContractRawList list)
    //    {
    //        Add(list);
    //    }

    //    public void Add(ContractRawList list)
    //    {
    //        if (this.Contracts == null)
    //        {
    //            lock (check)
    //            {
    //                if (this.Contracts == null)
    //                {
    //                    this.Contracts = new List<ContractRawList>();
    //                }
    //                else
    //                {
    //                    this.Contracts.Add(list);
    //                }
    //            }
    //        }
    //        else
    //        {
    //            this.Contracts.Add(list);
    //        }
    //    }


    //    public List<ContractRawList> Contracts { get; private set; }
    //}
   


    public class ContractsRawOfBill : IEnumerable<List<string>>
    {
        private ContractRawListNode first;

        public void Add(List<string> list)
        {
            if (this.first == null)
                this.first = new ContractRawListNode
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
        public ServicesOfBill() { }

        public ServicesOfBill(List<string> source) : base(source) { }

        public ServicesOfBill(List<ServiceOfBill> list)
        { Output = list; }

        public ServicesOfBill(ServicesOfBill services)
        { Output =services.Output; }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach (var s in Output)
            {
                sb.AppendLine(s.ToString());
            }

            return $"{sb.ToString()}\n";
        }
    }


    public class DetalizationOfContractOfBill : AbstractPartOfContractDetalization<StringOfDetalizationOfContractOfBill>//, IParseable
    {
        public DetalizationOfContractOfBill() { }
        public DetalizationOfContractOfBill(DetalizationOfContractOfBill detalization)
        {
            Output = detalization.Output;
        }

        public DetalizationOfContractOfBill(List<string> source) : base(source) { }

        public DetalizationOfContractOfBill(List<StringOfDetalizationOfContractOfBill> list)
        {
            Output = list;
        }

        public override string ToString()
        {
            StringBuilder sb = new StringBuilder();
            foreach(var s in Output)
            {
                sb.AppendLine(s.ToString());
            }

            return $"{sb.ToString()}\n";
        }
    }


    public class HeaderOfContractOfBill
    {
        public HeaderOfContractOfBill() { }
        public HeaderOfContractOfBill(ContractOfBill contract)
        {
            ContractId = contract.Header.ContractId;
            MobileNumber = contract.Header.MobileNumber;
            TarifPackage = contract.Header.TarifPackage;
        }

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

        public HeaderOfContractOfBill(List<string> source) { Source = source; }

        public List<string> Source { get; private set; }

        public string ContractId { get; private set; }

        public string MobileNumber { get; private set; }

        public string TarifPackage { get; private set; }

        public override string ToString()
        {
            return $"{ContractId}\t{MobileNumber}\t{TarifPackage}"; ;
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
            return $"{ServiceName}\t{NumberTarget}\t{Date}\t{Time}\t{DurationA}\t{DurationB}\t{Cost}";
        }
    }


    public class ServiceOfBill
    {
        public string Name { get; set; }

        public double Amount { get; set; }

        public bool IsMain { get; set; }

        public ServiceOfBill(string name, double amount, bool isMain=false)
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
