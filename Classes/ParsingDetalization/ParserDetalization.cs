using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
    public class ParserDetalization
    {
        List<ContractOfBill> contracts;
        List<string> contractRaw;
        List<string> billDetalizationList;
        string[] parsers;
        string startOfContract, stopParsing;
        public delegate void Status(object sender, TextEventArgs e);
        public event Status status;


        public ParserDetalization() { }

        public ParserDetalization(List<string> billDetalizationList, string[] parsers, string startOfContract, string stopParsing)
        {
            this.billDetalizationList = billDetalizationList;
            this.parsers = parsers;
            this.startOfContract = startOfContract;
            this.stopParsing = stopParsing;
        }

        public void SplitWholeBillToSeparatedContracts()
        {
            //Raw Contract data
            contractRaw = new List<string>(); //the whole list of contract detalization
            contracts = new List<ContractOfBill>();//List of Contracts with splited on separated parts

            foreach (var row in billDetalizationList)
            {
                //contract's Header
                if (row.StartsWith(startOfContract))
                {
                    if (contractRaw?.Count > 1) //if a contract already has all strings
                    {
                        contracts.Add(
                            new ContractOfBill(
                                contractRaw.SplitContractToSeparatedMainParts(parsers)
                                )
                            ); //create new contract with collected  of the List 'contractRaw' 
                    }

                    //Start new the List 'contractRaw'  and add first row 
                    contractRaw = new List<string> { row };
                }
                else if (row.StartsWith(stopParsing)) //After this parameter have no any Contract
                {
                    contracts.Add(new ContractOfBill(contractRaw)); //write last prepared contract

                    break;                  //After this parameter have no any Contract
                }
                else
                {
                    contractRaw.Add(row);   // add rows to created new the List 'contractRaw'  and add first row 
                }
            }
        }

        public List<ContractOfBill> GetContracts()
        {
            return contracts;
        }
    }


    public static class ParserDetalizationExtensions
    {
        public static ContractOfBill SplitContractToSeparatedMainParts(this List<string> theWholeContract, string[] parsers)
       //public static ContractOfBill SplitContractToSeparatedMainParts(this ContractOfBill contractOfBill, string[] parsers)
        {
            HeaderOfContractOfBill header = null;
            ServicesOfContractOfBill services = null;
            DetalizationOfContractOfBill billDetalizationList = null;

            List<string> partOfContract = new List<string>();

            if (theWholeContract?.Count > 0)
            {
                foreach (var row in theWholeContract)
                {
                    partOfContract.Add(row);

                    if (row.StartsWith(parsers[1])) // Start Header // "Контракт №"
                    {
                        partOfContract = new List<string>();
                        partOfContract.Add(row);
                    }
                    else if (row.StartsWith(parsers[3])) // Stop Header // "Ціновий Пакет" or "Тарифний Пакет"
                    {
                        header = new HeaderOfContractOfBill(partOfContract);

                        partOfContract = new List<string>(); //start part of Services
                    }
                    else if (row.StartsWith(parsers[7]))// Stop Services // @"ЗАГАЛОМ ЗА КОНТРАКТОМ (БЕЗ ПДВ ТА ПФ)"
                    {
                        partOfContract.ParseServicesOfContractOfBill(parsers);
                        services = new ServicesOfContractOfBill(partOfContract);
                        partOfContract = new List<string>(); //start part of Detalization
                    }
                }

                billDetalizationList = new DetalizationOfContractOfBill(partOfContract);
            }

            return new ContractOfBill(header, services, billDetalizationList);  //contractOfBill;
        }

        /// <summary>
        /// in source List: 
        /// first string looks like ' Контракт № 395409092966  Моб.номер: 380500251894' 
        /// second line - 'Ціновий Пакет: RED Business M'
        /// </summary>
        /// <param name="contractOfBill">wait ContractOfBill.Header.Source is not empty</param>
        /// <param name="parsers"></param>
        /// <returns></returns>
        public static ContractOfBill ParseHeaderOfContractOfBill(this ContractOfBill contractOfBill, string[] parsers)
        {
            ContractOfBill parsedContract = new ContractOfBill(contractOfBill); ;
            List<string> list = parsedContract.Header.Source;

            if (!(list?.Count > 0))
            { return null; }

            string contractId = "", mobileNumber = "", tarifPackage = "", tempRow;

            foreach (var rawData in list)
            {
                // if (!(rawData?.Length > 0))
                //     continue;

                if (rawData.Contains(parsers[1]))           //"Контракт №"  //Raw data = Контракт № 395409092966  Моб.номер: 380500251894 
                {

                    //\.{10,11}\s\d{11,12}\s{1,2}\.{15,16}\s\d{11,12}
                    //look for Contract's ID
                    contractId = System.Text.RegularExpressions.Regex.Split(rawData.Substring(rawData.IndexOf('№') + 1).Trim(), " ")[0].Trim();

                    //look for Contract's Mobile number
                    tempRow = rawData.Substring(rawData.IndexOf(':') + 1).Trim();
                    //set format number like '+380...'
                    if (tempRow.StartsWith("+"))
                    { mobileNumber = tempRow; }
                    else
                    { mobileNumber = "+" + tempRow; }
                }
                else if (rawData.Contains(parsers[3]))  //@"Ціновий Пакет" //Raw data = Ціновий Пакет: RED Business M
                {
                    tarifPackage = System.Text.RegularExpressions.Regex.Split(rawData.Substring(rawData.IndexOf(':') + 1).Trim(), " ")[0].Trim();
                }
            }

            parsedContract.Header = new HeaderOfContractOfBill(contractId, mobileNumber, tarifPackage);

            return parsedContract;
        }



        /// <summary>
        /// wait List<string> where string likes 'ВАРТІСТЬ ПАКЕТА/ЩОМІСЯЧНА ПЛАТА:  . . . . . . . . . . . . . . . . . . . .     0.0000  141.1760  141.1760'
        /// </summary>
        /// <param name="contractOfBill">wait ContractOfBill.ServicesOfContract.Source is not empty</param>
        /// <param name="parsers"></param>
        /// <returns></returns>
        public static ServicesOfContractOfBill ParseServicesOfContractOfBill(this List<string> list, string[] parsers)
        {
            if (!(list?.Count > 0))
            { return null; }

            List<ServiceOfBill> services = new List<ServiceOfBill>();
            double cost = 0;

            foreach (var rawData in list)
            {
                foreach (string parser in parsers)
                {
                    if (rawData.Contains(parser) && rawData.Length > 96)
                    {
                        string parsed = rawData.Substring(rawData.LastIndexOf(' '))?.Trim(); //rawData likes 'ВАРТІСТЬ ПАКЕТА/ЩОМІСЯЧНА ПЛАТА:  . . . . . . . . . . . . . . . . . . . .     0.0000  141.1760  141.1760'
                        if (double.TryParse(parsed, out cost))
                        {
                            services.Add(new ServiceOfBill(parser, cost));
                        }

                        break;
                    }
                }
            }

            return new ServicesOfContractOfBill(services);
        }

        /// <summary>
        /// Return new DetalizationOfContractOfBill which contained parsed detalization
        /// </summary>
        /// <param name="contractOfBill"></param>
        /// <param name="parsers"></param>
        /// <returns></returns>
        public static DetalizationOfContractOfBill ParseDetalizationOfContractOfBill(this ContractOfBill contractOfBill, string[] parsers)
        {
            List<string> list = contractOfBill?.DetalizationOfContract?.Source;
            if (!(list?.Count > 0)) { return null; }

            List<StringOfDetalizationOfContractOfBill> strings = new List<StringOfDetalizationOfContractOfBill>();

            foreach (var stringOfDetalization in list)
            {
                strings.Add(stringOfDetalization.ParseStringOfDetalizationOfContractOfBill());
            }

            return new DetalizationOfContractOfBill(strings);
        }

        /// <summary>
        /// Return StringOfDetalizationOfContractOfBill which concluded a Parsed String Of Detalization Of Contract Of Bill
        /// </summary>
        /// <param name="DetalizationString"> The whole parsed parameters of string of detalization</param>
        /// <returns></returns>
        public static StringOfDetalizationOfContractOfBill ParseStringOfDetalizationOfContractOfBill(this string DetalizationString)
        {
            if (DetalizationString?.Length < 100)
            { return null; }

            // status?.Invoke(this, new TextEventArgs(DetalizationString));

            return new StringOfDetalizationOfContractOfBill
            {
                ServiceName = DetalizationString?.Substring(0, 38)?.Trim() ?? "",
                NumberTarget = DetalizationString?.Substring(38, 13)?.Trim() ?? "",
                Date = DetalizationString?.Substring(52, 10)?.Trim() ?? "",
                Time = DetalizationString?.Substring(65, 8)?.Trim() ?? "",
                DurationA = DetalizationString?.Substring(74, 9)?.Trim() ?? "",
                DurationB = DetalizationString?.Substring(84, 9)?.Trim() ?? "",
                Cost = DetalizationString?.Substring(95)?.Trim() ?? ""
            }; 
        }
    }
}
