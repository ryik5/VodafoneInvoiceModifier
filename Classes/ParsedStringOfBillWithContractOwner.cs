using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MobileNumbersDetailizationReportGenerator
{
    public class ParsedStringOfBill
    {
        public string contract { get; set; }//= "";
        public string numberOwner { get; set; }//= "";
        public string serviceName { get; set; }//= "";
        public string numberTarget { get; set; }//= "";
        public string date { get; set; }//= "";
        public string time { get; set; }//= "";
        public string durationA { get; set; }//= "";
        public string durationB { get; set; }//= "";
        public string cost { get; set; }//= "";

        public string fio { get; set; }// = "";
        public string nav { get; set; }//= "";
        public string department { get; set; }//= "";
    }

    /// <summary>
    /// inline string of detalization Bill must be a char position
    /// /                        
    /// <param name="detalizationString">
    /// 1-39	наименование услуги /
    /// 40-52	номер(целевой) /
    /// 53-63	дата /
    /// 66-74	время /
    /// 75-84	длительность /
    /// 85-95	учтенная длительность оператором(для биллинга) /
    /// 96-106	стоимость /
    ///</param>>
    /// </summary>
    public class ParsingStringDetalizationOfBill
    {
        string _detalizationString;
        ParsedStringOfBill parsedString;

        public ParsingStringDetalizationOfBill() { }

        public ParsingStringDetalizationOfBill(string detalizationString)
        {
            SetString(detalizationString);
        }

        public void SetString(string detalizationString)
        {
            _detalizationString = detalizationString;
        }

        public ParsedStringOfBill ParseString()
        {
            if (_detalizationString == null)
                return null;

            parsedString = new ParsedStringOfBill
            {
                serviceName = _detalizationString?.Substring(0, 38)?.Trim(),
                numberTarget = _detalizationString?.Substring(38, 13)?.Trim(),
                date = _detalizationString?.Substring(52, 10)?.Trim(),
                time = _detalizationString?.Substring(65, 8)?.Trim(),
                durationA = _detalizationString?.Substring(74, 9)?.Trim(),
                durationB = _detalizationString?.Substring(84, 9)?.Trim(),
                cost = _detalizationString?.Substring(95)?.Trim()
            };

            return parsedString;
        }
    }
}
