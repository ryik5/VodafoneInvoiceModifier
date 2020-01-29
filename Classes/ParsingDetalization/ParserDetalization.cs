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
        ContractsRawOfBill wholeContractsRaw;

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

            status?.Invoke(this, new TextEventArgs("исходный список, строк: "+ this.billDetalizationList.Count.ToString()));

        }

        public void SplitWholeBillToSeparatedContracts()
        {
            //second variant
            //Raw Contract data
            contractRaw = new List<string>(); //the whole list of contract detalization
            contracts = new List<ContractOfBill>();//List of Contracts with splited on separated parts

            //first variant
            wholeContractsRaw = new ContractsRawOfBill();
            //  ContractRawList ContractRaw = new ContractRawList();

            foreach (var row in billDetalizationList)
            {
                //contract's Header
                if (row.StartsWith(startOfContract))
                {
                    if (contractRaw?.Count > 1) //if a contract already has all strings
                    {
                        //first variant
                        wholeContractsRaw.Add(contractRaw);


                        //second variant - check
                        //delelete?
                        contracts.Add(
                            new ContractOfBill(
                                contractRaw.SplitWholeContractToSeparatedMainParts(parsers)
                                )
                            ); //create new contract with collected  of the List 'contractRaw' 
                    }

                    //Start new the List 'contractRaw'  and add first row 
                    contractRaw = new List<string> { row };
                }
                else if (row.StartsWith(stopParsing)) //After this parameter have no any Contract
                {
                    //first variant
                    //first variant
                    wholeContractsRaw.Add(contractRaw);


                    //second variant - check
                    //delelete?
                    contracts.Add(new ContractOfBill(contractRaw)); //write last prepared contract

                    break;                  //After this parameter have no any Contract
                }
                else
                {
                    contractRaw.Add(row);   // add rows to created new the List 'contractRaw'  and add first row 
                }
            }
            status?.Invoke(this, new TextEventArgs("контрактов с сырыми данными, шт: " + this.contracts.Count.ToString()));

        }

        public List<ContractOfBill> ParseContracts()
        {
           // var contracts = wholeContractsRaw.Select(s => s.SplitWholeContractToSeparatedMainParts(parsers));
            
            List<ContractOfBill> result = new List<ContractOfBill>();

            foreach (var contractRaw in wholeContractsRaw)
            {
                ContractOfBill contract = new ContractOfBill(contractRaw.SplitWholeContractToSeparatedMainParts(parsers));

                result.Add(contract);
            }

            status?.Invoke(this, new TextEventArgs("Распарсеных контрактов, шт: " + result.Count.ToString()));

            return result;
        }



    }


    public static class ParserDetalizationExtensions
    {
        public static ContractOfBill SplitWholeContractToSeparatedMainParts(this List<string> theWholeContract, string[] parsers)
        {
            HeaderOfContractOfBill header = null;
            ServicesOfBill services = null;
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
                        header = new HeaderOfContractOfBill(partOfContract.ParseHeaderOfContractOfBill(parsers));

                        partOfContract = new List<string>(); //start part of Services
                    }
                    else if (row.StartsWith(parsers[7]))// Stop Services // @"ЗАГАЛОМ ЗА КОНТРАКТОМ (БЕЗ ПДВ ТА ПФ)"
                    {
                        services = new ServicesOfBill(partOfContract.ParseServicesOfBill(parsers));

                        partOfContract = new List<string>(); //start part of Detalization
                    }
                }

                billDetalizationList = new DetalizationOfContractOfBill(partOfContract.ParseDetalizationOfContractOfBill());
            }

            return new ContractOfBill(header, services, billDetalizationList);  //contractOfBill;
        }

        /// <summary>
        /// in source List: 
        /// first string should look like 'Контракт № 395409092966  Моб.номер: 380500251894' 
        /// second line - 'Ціновий Пакет: RED Business M'
        /// </summary>
        /// <param name="list"> wait List<string> as the first 2 strings of the whole Contract</param>
        /// <param name="parsers"></param>
        /// <returns></returns>
        public static HeaderOfContractOfBill ParseHeaderOfContractOfBill(this List<string> list, string[] parsers)
        {
            HeaderOfContractOfBill header = new HeaderOfContractOfBill();

            if (!(list?.Count > 0))
            { return header; }

            string contractId = "", mobileNumber = "", tarifPackage = "";

            foreach (var rawData in list)
            {
                // if (!(rawData?.Length > 0))
                //     continue;

                if (rawData.Contains(parsers[1]))           //"Контракт №"  //Raw data = Контракт № 395409092966  Моб.номер: 380500251894 
                {
                    //look for Contract's ID
                    contractId = GetContractId(rawData);

                    //look for Contract's Mobile number
                    mobileNumber = GetMobileNumber(rawData);
                }
                else if (rawData.Contains(parsers[3]))  //@"Ціновий Пакет" //Raw data = Ціновий Пакет: RED Business M
                {
                    tarifPackage = GetTarifPackage(rawData);
                }
            }
            header = new HeaderOfContractOfBill(contractId, mobileNumber, tarifPackage);

            return header;
        }

        public static string GetContractId(string data)
        {
            return System.Text.RegularExpressions.Regex.Split(data.Substring(data.IndexOf('№') + 1).Trim(), " ")[0].Trim();
        }

        public static string GetMobileNumber(string data)
        {
            string mobileNumber;
            string tempRow = data.Substring(data.IndexOf(':') + 1).Trim();
            //set format number like '+380...'
            if (tempRow.StartsWith("+"))
            { mobileNumber = tempRow; }
            else
            { mobileNumber = "+" + tempRow; }

            return mobileNumber;
        }
        
        public static string GetTarifPackage(string data)
        {
            return data.Substring(data.IndexOf(':') + 1).Trim();
        }


        /// <summary>
        /// wait List<string> where string likes 'ВАРТІСТЬ ПАКЕТА/ЩОМІСЯЧНА ПЛАТА:  . . . . . . . . . . . . . . . . . . . .     0.0000  141.1760  141.1760'
        /// </summary>
        /// <param name="contractOfBill">wait ContractOfBill.ServicesOfContract.Source is not empty</param>
        /// <param name="parsers"></param>
        /// <returns></returns>
        public static ServicesOfBill ParseServicesOfBill(this List<string> list, string[] parsers)
        {
            if (!(list?.Count > 0))
            { return null; }

            List<ServiceOfBill> services = new List<ServiceOfBill>();
            double cost = 0;
            string name;

            foreach (var rawData in list)
            {
              //  foreach (string parser in parsers)
                {
                    if (rawData.Length > 96) //rawData.Contains(parser) && 
                    {
                        cost = rawData.ParseCostOfServiceOfBill();
                        name = rawData.ParseNameOfServiceOfBill(':');

                        if (name != null)
                        {
                            services.Add(new ServiceOfBill(name, cost));
                        }
                        break;
                    }
                }
            }

            return new ServicesOfBill(services);
        }

        /// <summary>
        /// rawData likes 'ВАРТІСТЬ ПАКЕТА/ЩОМІСЯЧНА ПЛАТА:  . . . . . . . . . . . . . . . . . . . .     0.0000  141.1760  141.1760'
        /// it will be returned double '141.1760'
        /// </summary>
        /// <param name="rawString">likes 'ВАРТІСТЬ ПАКЕТА/ЩОМІСЯЧНА ПЛАТА:  . . . . . . . . . . . . . . . . . . . .     0.0000  141.1760  141.1760'</param>
        public static double ParseCostOfServiceOfBill(this string rawString)
        {
            if (!rawString.Contains(' '))
                return 0;

            double cost = 0;
            string parsed = rawString.Substring(rawString.LastIndexOf(' '))?.Trim();
            if (double.TryParse(parsed, out cost))
            {
                return cost;
            }
            else
            {
                return 0;
            }
        }
        public static string ParseNameOfServiceOfBill(this string rawString, char parser)
        {
            if (!rawString.Contains(parser))
                return null;

            string parsed = rawString.Substring(0,rawString.IndexOf(parser))?.Trim();
            return parsed;
        }


        /// <summary>
        /// Return new DetalizationOfContractOfBill which contained parsed detalization
        /// </summary>
        /// <param name="contractOfBill"></param>
        /// <param name="parsers"></param>
        /// <returns></returns>
        public static DetalizationOfContractOfBill ParseDetalizationOfContractOfBill(this List<string> list)
        {
            DetalizationOfContractOfBill detalization = new DetalizationOfContractOfBill();
            if (!(list?.Count > 0))
            { return detalization; }

            List<StringOfDetalizationOfContractOfBill> detalizationStrings = new List<StringOfDetalizationOfContractOfBill>();

            foreach (var stringOfDetalization in list)
            {
                detalizationStrings.Add(
                    stringOfDetalization.ParseStringOfDetalizationOfContractOfBill() //each string is detalized to object StringOfDetalizationOfContractOfBill 
                    );
            }

            detalization = new DetalizationOfContractOfBill(detalizationStrings);

            return detalization;
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
