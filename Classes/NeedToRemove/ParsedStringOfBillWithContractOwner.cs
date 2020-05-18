namespace MobileNumbersDetailizationReportGenerator
{


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

        string DetalizationString { get; set; }

        ParsedContractOfBill ParsedString { get; set; }

        public ParsingStringDetalizationOfBill() { }

        public ParsingStringDetalizationOfBill(string detalizationString, ParsedContractOfBill parsedString)
        {
            DetalizationString = detalizationString;
            ParsedString = parsedString;
        }


        /// <summary>
        /// if string of Detalization has correct's length (from 95 to) it will return true
        /// </summary>
        /// <returns></returns>
        public bool ParseStringOfBodyOfContractOfBill()
        {
            if (!(DetalizationString?.Length>0)|| DetalizationString?.Length < 102)
                return false;

            if (ParsedString == null)
            { ParsedString = new ParsedContractOfBill(); }

            status?.Invoke(this, new TextEventArgs(DetalizationString));

            ParsedString.ServiceName = DetalizationString?.Substring(0, 38)?.Trim()??"";
            ParsedString.NumberTarget = DetalizationString?.Substring(38, 13)?.Trim() ?? "";
            ParsedString.Date = DetalizationString?.Substring(52, 10)?.Trim() ?? "";
            ParsedString.Time = DetalizationString?.Substring(65, 8)?.Trim() ?? "";
            ParsedString.DurationA = DetalizationString?.Substring(74, 9)?.Trim() ?? "";
            ParsedString.DurationB = DetalizationString?.Substring(84, 9)?.Trim() ?? "";
            ParsedString.Cost = DetalizationString?.Substring(95)?.Trim() ?? "";
          
            return true;
        }

        public ParsedContractOfBill Get()
        { return ParsedString; }
    }
}
