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
        string beginningOfFirstLineOfContract, stopParsing;
        public delegate void Status(object sender, TextEventArgs e);
        public event Status Info;


        public ParserDetalization() { }

        public ParserDetalization(List<string> billDetalizationList, string[] parsers, string beginningOfFirstLineOfContract, string stopParsing)
        {
            this.billDetalizationList = billDetalizationList;
            this.parsers = parsers;
            this.beginningOfFirstLineOfContract = beginningOfFirstLineOfContract;
            this.stopParsing = stopParsing;

            Info?.Invoke(this, new TextEventArgs("исходный счет, строк: " + this.billDetalizationList?.Count.ToString()));

            Info?.Invoke(this, new TextEventArgs("парсеров: " + this.parsers?.Length.ToString()));
            Info?.Invoke(this, new TextEventArgs("старт-парсер: " + this.beginningOfFirstLineOfContract));
            Info?.Invoke(this, new TextEventArgs("стоп-парсер: " + this.stopParsing));
        }

        public ContractsRawOfBill SplitWholeBillToSeparatedContracts()
        {
            //second variant
            //Raw Contract data
            contractRaw = new List<string>(); //the whole list of contract detalization
            contracts = new List<ContractOfBill>();//List of Contracts with splited on separated parts
            bool headerOfBillFinished = false;
            //first variant
            ContractsRawOfBill wholeContractsRaw = new ContractsRawOfBill();

            foreach (var row in billDetalizationList)
            {

                //contract's Header
                if (row.StartsWith(beginningOfFirstLineOfContract))
                {
                    headerOfBillFinished = true;//Bill's header was finished

                    if (contractRaw?.Count > 2) //if a contract already has had strings
                    {
                        wholeContractsRaw.Add(contractRaw);
                    }

                    //Start new the List 'contractRaw'  and add first row 
                    contractRaw = new List<string> { row };
                }
                else if (row.StartsWith(stopParsing)) //After this parameter the bill has no contained contracts
                {
                    wholeContractsRaw.Add(contractRaw);

                    break;
                }
                else
                {
                    contractRaw.Add(row);   // add rows to created new the List 'contractRaw'  and add first row 
                }

                if (!headerOfBillFinished)
                {
                }
            }
            Info?.Invoke(this, new TextEventArgs("контрактов с сырыми данными, шт: " + wholeContractsRaw.Count().ToString()));

            return wholeContractsRaw;
        }

        public ServicesOfBill GetHeaderOfBill()
        {
            //second variant
            //Raw Contract data
            contractRaw = new List<string>(); //the whole list of contract detalization

            foreach (var row in billDetalizationList)
            {

                //contract's Header
                if (row.StartsWith(beginningOfFirstLineOfContract))
                {
                    break;
                }
                else
                {
                    contractRaw.Add(row);   // add rows to created new the List 'contractRaw'  and add first row 
                }
            }
            Info?.Invoke(this, new TextEventArgs("В шапке счета строк: " + contractRaw.Count().ToString()));

            ServicesOfBill header = new ServicesOfBill(contractRaw.ParseServicesOfBill());

            return header;
        }

        public List<ContractOfBill> ParseContracts(ContractsRawOfBill wholeContractsRaw)
        {
            List<ContractOfBill> result = new List<ContractOfBill>();

            foreach (var contractRaw in wholeContractsRaw)
            {
                //todo
                //new task for every parsing
                ContractOfBill contract = new ContractOfBill(contractRaw.SplitWholeContractToSeparatedMainParts(parsers));

                result.Add(contract);
            }

            Info?.Invoke(this, new TextEventArgs("Разделенных контрактов: " + result.Count.ToString()));

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
                        services = new ServicesOfBill(partOfContract.ParseServicesOfBill());

                        partOfContract = new List<string>(); //start part of Detalization
                    }
                }

                billDetalizationList = new DetalizationOfContractOfBill(partOfContract.ParseDetalizationOfContractOfBill());
            }
            if (!(header?.ContractId?.Length > 0))
                return null;

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
        public static ServicesOfBill ParseServicesOfBill(this List<string> list)
        {
            if (!(list?.Count > 0))
            { return null; }

            List<ServiceOfBill> services = new List<ServiceOfBill>();
            double cost;
            string name;
            bool isMain;

            foreach (var rawData in list)
            {
                if (rawData.Length > 96) //rawData.Contains(parser) && 
                {
                    cost = rawData.ParseCostOfServiceOfBill();

                    if (rawData.Contains(':'))
                    {
                        isMain = true;
                    }
                    else
                    {
                        isMain = false;
                    }

                    name = rawData.ParseNameOfServiceOfBill(':');

                    if (name != null)
                    {
                        services.Add(new ServiceOfBill(name, cost, isMain));
                    }
                }
            }

            return new ServicesOfBill(services);
        }

        /// <summary>
        /// rawData likes 'ВАРТІСТЬ ПАКЕТА/ЩОМІСЯЧНА ПЛАТА:  . . . . . . . . . . . . . . . . . . . .     0.0000  141.1760  141.1760'
        /// it will be returned double '141.1760'
        /// </summary>
        /// <param name="rawData">likes 'ВАРТІСТЬ ПАКЕТА/ЩОМІСЯЧНА ПЛАТА:  . . . . . . . . . . . . . . . . . . . .     0.0000  141.1760  141.1760'</param>
        public static double ParseCostOfServiceOfBill(this string rawData)
        {
            double cost = 0;
            string parsed = rawData.Substring(rawData.LastIndexOf(' '))?.Trim();
            if (double.TryParse(parsed, out cost))
            {
                return cost;
            }
            else
            {
                return 0;
            }
        }

        public static string ParseNameOfServiceOfBill(this string rawString, char parserInLineWithMainService)
        {
            if (rawString == null)
            { return null; }

            string parsed;
            if (rawString.Contains(parserInLineWithMainService)) //line with main service should contains 'parserInLineWithMainService'
            {
                parsed = rawString.Substring(0, rawString.IndexOf(parserInLineWithMainService))?.Trim();
            }
            else                  //every line with service contains - '. . . . . . . . '
            {
                parsed = rawString.Substring(0, rawString.IndexOf(". ."))?.Trim();
            }

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
            if (list == null || !(list?.Count > 0))
            { return new DetalizationOfContractOfBill(); }

            List<StringOfDetalizationOfContractOfBill> detalizationStrings = new List<StringOfDetalizationOfContractOfBill>();

            foreach (var stringOfDetalization in list)
            {
                detalizationStrings.Add(
                    stringOfDetalization.ParseStringOfDetalizationOfContractOfBill() //each string is detalized to object StringOfDetalizationOfContractOfBill 
                    );
            }

            DetalizationOfContractOfBill detalization = new DetalizationOfContractOfBill(detalizationStrings);

            return detalization;
        }

        /// <summary>
        /// Return StringOfDetalizationOfContractOfBill which concluded a Parsed String Of Detalization Of Contract Of Bill
        /// </summary>
        /// <param name="DetalizationString"> The whole parsed parameters of string of detalization</param>
        /// <returns></returns>
        public static StringOfDetalizationOfContractOfBill ParseStringOfDetalizationOfContractOfBill(this string DetalizationString)
        {
            if (DetalizationString == null || DetalizationString?.Length < 100)
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
