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

        List<string> detalization;
        List<string> result;
        string[] parsers;
        string parametrStart, pStop;

        public delegate void Status(object sender, TextEventArgs e);
        public event Status status;


        public ParserDetalization() { }

        public ParserDetalization(List<string> billDetalizationList, string[] parsers,
            string startOfContract,
            string stopParsing)
        {
            detalization = billDetalizationList;
            this.parsers = parsers;
        }

        public void SplitBillToContracts()
        {
            bool headerCorrect = false;
            bool headerFinished = false;
            bool firstStringAtDetalizationContract = false;

            //Raw Contract data
            contractRaw = new List<string>();
            contracts = new List<ContractOfBill>();

            foreach (var row in detalization)
            {
                //contract's Header
                if (row.StartsWith(parametrStart))
                {
                    if (contractRaw?.Count > 1) //if a contract already has all strings
                    {
                        contracts.Add(new ContractOfBill { Source = contractRaw });
                    }

                    //Start new Making contract
                    contractRaw = new List<string> { row }; 
                }
                else if (row.StartsWith(pStop)) //After this parameter have no any Contract
                {
                    contracts.Add(new ContractOfBill { Source = contractRaw });
                    break;
                }
                else
                {
                    contractRaw.Add(row);
                }
            }
          //      @"Ціновий Пакет",                                      //3     //name of tarif package
        //    @"ЗАГАЛОМ ЗА КОНТРАКТОМ (БЕЗ ПДВ ТА ПФ)",           //7     //total without tax and pf


                //switch (contract)
                //{
                //    case StringOfDetalizationsOfContract.ContractIdentification:
                //        {
                //            //Parse the Contract's first row Header of detalization
                //            headerCorrect = detalization.ParseHeaderContract(row);
                //            parsedHeaderContract = detalization.Get();
                //            //    ("Проверьте правильность выбора файла с контрактами с детализацией разговоров!" + Environment.NewLine +
                //            //    "Возможно поменялся формат." + Environment.NewLine +
                //            //    "Правильный формат первых строк с новым контрактом:" + Environment.NewLine +
                //            //    @"Контракт №" + " 000000000  Моб.номер: 380000000000" + Environment.NewLine +
                //            //    "Ціновий Пакет: название_пакета" + Environment.NewLine + "далее - детализацией разговоров контракта" + Environment.NewLine +
                //            //    "В данном случае строка с началом разбираемого контракта имеет форму:" + Environment.NewLine +
                //            //    row + Environment.NewLine + "Ошибка: " + err.ToString());
                //            break;
                //        }
                //    case StringOfDetalizationsOfContract.Header: //If Contract was started but the its header isn't finished yet
                //        {
                //            //Parse start of Contract's Header of detalization
                //            //it is contract's header parsing
                //            break;
                //        }
                //    case StringOfDetalizationsOfContract.Body:  //If Contract was started, its header finished but detalization isn't finished yet
                //        {
                //            //it is contract's body detalization parsing
                //            detalization = new ParsingStringDetalizationOfBill(row, parsedHeaderContract);
                //            detalization.ParseRowFromTheBodyDetalizationContract();
                //            parsedBodyContract = detalization.Get();
                //            parsedList.Add(parsedBodyContract);
                //            break;
                //        }
                //    case StringOfDetalizationsOfContract.Stop:
                //        break;
                //}
                       
        }

        public void SplitContractToParts(ContractOfBill contract, string[] parsers)
        {
            HeaderOfContractOfBill header = new HeaderOfContractOfBill();
            ServicesOfContractOfBill services = new ServicesOfContractOfBill();
            DetalizationOfContractOfBill detalization = new DetalizationOfContractOfBill();

            List<string> partOfContract = new List<string>();
            if (contract?.Source?.Count > 0)
            {
                foreach (var row in contract.Source)
                {
                    partOfContract.Add(row);
                    if (row.StartsWith(parsers[3]))
                    {

                    }
                }
            }
        }

        public void Get()
        {

        }
    }


}
