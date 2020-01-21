using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
    public class ParserDetalization
    {

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

        public void Parse()
        {
            bool headerCorrect = false;
            bool headerFinished = false;
            bool firstStringAtDetalizationContract = false;

            //Raw Contract data
            List<string> contractRaw = new List<string>();
            List<ContractOfBill> contracts = new List<ContractOfBill>();

            foreach (var row in detalization)
            {
                if (contractRaw?.Count > 1)
                    contracts.Add(contractRaw);

                //contract's Header
                if (row.StartsWith(parametrStart))
                {

                    //run new block
                    contractRaw = new List<string> { row };
                }
                else if (row.StartsWith(pStop))
                {

                    break;
                }
                else
                {
                    contractRaw.Add(row);
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
        }

        public void Get()
        {

        }
    }


}
