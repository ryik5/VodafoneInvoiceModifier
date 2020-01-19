using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
    public class ParserDetalization : IDetalizationParseable
    {
        List<string> detalization;
        List<string> result;
        string[] parsers;
            List<ParsedContractOfBill> parsedList = new List<ParsedContractOfBill>();
            ParsingStringDetalizationOfBill parsed = new ParsingStringDetalizationOfBill();
            ParsedContractOfBill parsedBodyContract = new ParsedContractOfBill();
            ParsedContractOfBill parsedHeaderContract = new ParsedContractOfBill();
            StringOfDetalizationsOfContract contract = new StringOfDetalizationsOfContract();
        string parametrStart, pStop;

        public delegate void Status(object sender, TextEventArgs e);
        public event Status status;


        public ParserDetalization() { }
        
        public ParserDetalization(List<string> billDetalizationList, string[] parsers, string parametrStart, string pStop)
        {
            detalization = billDetalizationList;
            this.parsers = parsers;
        }

        public void Parse()
        {
            bool headerCorrect = false;
            bool headerFinished = false;
            bool firstStringAtDetalizationContract = false;

            foreach (var row in detalization)
            {

                //contract's Header
                if (row.StartsWith(parametrStart))
                {
                    contract = StringOfDetalizationsOfContract.FirstRow;
                }
                else if (row.StartsWith(parsers[3]))
                {
                    contract = StringOfDetalizationsOfContract.Header;
                    continue;
                }
                else if (row.StartsWith(pStop))
                {
                    contract = StringOfDetalizationsOfContract.Stop;
                    break;
                }
                else if (row.StartsWith(parsers[7]))
                {
                    contract = StringOfDetalizationsOfContract.Body;
                    //строку обработать
                    continue;
                }
                /
         @"Ціновий Пакет",                                      //3     //name of tarif package
            @"ЗАГАЛОМ ЗА КОНТРАКТОМ (БЕЗ ПДВ ТА ПФ)",           //7     //total without tax and pf


                switch (contract)
                {
                    case StringOfDetalizationsOfContract.FirstRow:
                        {
                            //Parse the Contract's first row Header of detalization
                            headerCorrect = detalization.ParseHeaderContract(row);
                            parsedHeaderContract = detalization.Get();
                            //    ("Проверьте правильность выбора файла с контрактами с детализацией разговоров!" + Environment.NewLine +
                            //    "Возможно поменялся формат." + Environment.NewLine +
                            //    "Правильный формат первых строк с новым контрактом:" + Environment.NewLine +
                            //    @"Контракт №" + " 000000000  Моб.номер: 380000000000" + Environment.NewLine +
                            //    "Ціновий Пакет: название_пакета" + Environment.NewLine + "далее - детализацией разговоров контракта" + Environment.NewLine +
                            //    "В данном случае строка с началом разбираемого контракта имеет форму:" + Environment.NewLine +
                            //    row + Environment.NewLine + "Ошибка: " + err.ToString());
                            break;
                        }
                    case StringOfDetalizationsOfContract.Header: //If Contract was started but the its header isn't finished yet
                        {
                            //Parse start of Contract's Header of detalization
                            //it is contract's header parsing
                            break;
                        }
                    case StringOfDetalizationsOfContract.Body:  //If Contract was started, its header finished but detalization isn't finished yet
                        {
                            //it is contract's body detalization parsing
                            detalization = new ParsingStringDetalizationOfBill(row, parsedHeaderContract);
                            detalization.ParseRowFromTheBodyDetalizationContract();
                            parsedBodyContract = detalization.Get();
                            parsedList.Add(parsedBodyContract);
                            break;
                        }
                    case StringOfDetalizationsOfContract.Stop:
                        break;
                }
            }
        }
        
        public void Get()
        {

        }
    }

   public interface IDetalizationParseable
    {
        void Parse();
        void Get();
    }
}
