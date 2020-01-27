using MobileNumbersDetailizationReportGenerator;
using NUnit.Framework;
using System.Collections.Generic;

namespace NUnitTestProject
{
    [TestFixture]
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void TestParsingHeaderContracts()
        {
            //Arrange                
            List<string> inputed = new List<string>();
            inputed.Add("Контракт № 395409092966  Моб.номер: 380500251894");
            inputed.Add("Ціновий Пакет: RED Business M ");
            
            string[] parsers = new string[] {
                @"Владелец",
                @"Контракт №",
                @"Моб.номер",
                @"Тарифний Пакет"
            };


            //Act
            var result = ParserDetalizationExtensions.ParseHeaderOfContractOfBill(inputed, parsers);


            //Assert
            Assert.AreEqual(200, result);
        }

        [Test]
        public void TestToInternetTrafic_200_Mb_Wait_200()
        {
            string text = "200 Mb";
            var result = WinFormsExtensions.ToInternetTrafic(text, "Mb");

            Assert.AreEqual(200, result);
        }

        [Test()]
        public void TestToInternetTrafic_10_Kb_Wait_10()
        {
            //Arrange
            string text = "10 Kb";

            //Act
            var result = WinFormsExtensions.ToInternetTrafic(text, "Kb");

            //Assert
            Assert.AreEqual(result, 10);
        }
    }
}