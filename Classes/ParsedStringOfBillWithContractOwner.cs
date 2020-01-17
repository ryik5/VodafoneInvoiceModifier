
using System.Text.RegularExpressions;

namespace MobileNumbersDetailizationReportGenerator
{
    public class ParsedContractOfBill
    {
        public string contract { get; set; }
        public string numberOwner { get; set; }
        public string serviceName { get; set; }
        public string numberTarget { get; set; }
        public string date { get; set; }
        public string time { get; set; }
        public string durationA { get; set; }
        public string durationB { get; set; }
        public string cost { get; set; }

        public string fio { get; set; }
        public string nav { get; set; }
        public string department { get; set; }
    }

    /// <summary>
    /// inline string of detalization Bill must be a char position
    /// </summary>                        
    /// <param name="detalizationString">
    /// 1-39	наименование услуги /
    /// 40-52	номер(целевой) /
    /// 53-63	дата /
    /// 66-74	время /
    /// 75-84	длительность /
    /// 85-95	учтенная длительность оператором(для биллинга) /
    /// 96-106	стоимость /
    ///</param>>
    public class ParsingStringDetalizationOfBill
    {
        public delegate void Status(object sender, TextEventArgs e);

        public event Status status;

        string detalisation;
        string Detalization { get { return detalisation; } set { detalisation = value; } }

        ParsedContractOfBill ParsedString { get; set; }

        public ParsingStringDetalizationOfBill() { }

        public ParsingStringDetalizationOfBill(string detalizationString)
        { Detalization = detalizationString; }

        public ParsingStringDetalizationOfBill(string detalizationString, ParsedContractOfBill parsedString)
        {
            Detalization = detalizationString;
            ParsedString = parsedString;
        }

        public ParsingStringDetalizationOfBill(ParsedContractOfBill parsedString)
        { ParsedString = parsedString; }

        public bool Parse()
        {
            if (!(Detalization?.Length>0)|| Detalization?.Length < 102)
                return false;

            if (ParsedString == null)
            { ParsedString = new ParsedContractOfBill(); }

            status?.Invoke(this, new TextEventArgs(Detalization));

            ParsedString.serviceName = Detalization?.Substring(0, 38)?.Trim()??"";
            ParsedString.numberTarget = Detalization?.Substring(38, 13)?.Trim() ?? "";
            ParsedString.date = Detalization?.Substring(52, 10)?.Trim() ?? "";
            ParsedString.time = Detalization?.Substring(65, 8)?.Trim() ?? "";
            ParsedString.durationA = Detalization?.Substring(74, 9)?.Trim() ?? "";
            ParsedString.durationB = Detalization?.Substring(84, 9)?.Trim() ?? "";
            ParsedString.cost = Detalization?.Substring(95)?.Trim() ?? "";
          
            return true;
        }

        public bool ParseHeaderContract(string headerContract)
        {
            if (!(headerContract?.Trim()?.Length > 0))
                return false;

            ParsedString = new ParsedContractOfBill
            {
                contract = Regex.Split(headerContract.Substring(headerContract.IndexOf('№') + 1).Trim(), " ")[0].Trim()
            };

            string tempRow = headerContract.Substring(headerContract.IndexOf(':') + 1).Trim();
 
            //set format number like '+380...'
            if (tempRow.StartsWith("+"))
            { ParsedString.numberOwner = tempRow; }
            else
            { ParsedString.numberOwner = "+" + tempRow; }
            
            //   "Проверьте правильность выбора файла с контрактами с детализацией разговоров!" + Environment.NewLine +
            //    "Возможно поменялся формат." + Environment.NewLine +
            //    "Правильный формат первых строк с новым контрактом:" + Environment.NewLine +
            //    NUMBER_OF_CONTRACT + " 000000000  Моб.номер: 380000000000" + Environment.NewLine +
            //    "Ціновий Пакет: название_пакета" + Environment.NewLine + "далее - детализацией разговоров контракта" + Environment.NewLine +
            //    "В данном случае строка с началом разбираемого контракта имеет форму:" + Environment.NewLine +
            //    row + Environment.NewLine + "Ошибка: " + err.ToString()
            
            return true;
        }

        public ParsedContractOfBill Get()
        { return ParsedString; }
    }
}
