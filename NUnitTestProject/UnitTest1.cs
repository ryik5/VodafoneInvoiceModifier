using MobileNumbersDetailizationReportGenerator;
using NUnit.Framework;
using System.Collections.Generic;

namespace NUnitTestProject
{
 //   [TestFixture]
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
        }




        [Test]
        public void TestParseCostOfServiceOfBill()
        {
            //Arrange
            string text = @"ВАРТІСТЬ ПАКЕТА/ЩОМІСЯЧНА ПЛАТА:  . . . . . . . . . . . . . . . . . . . .     0.0000  141.1760  141.1760";

            //Act
            var result = ParserDetalizationExtensions.ParseCostOfServiceOfBill(text);

            //Assert
            Assert.AreEqual(141.1760, result);
        }

        [Test]
        public void TestParseNameOfServiceOfBill()
        {
            //Arrange
            string text = @"ВАРТІСТЬ ПАКЕТА/ЩОМІСЯЧНА ПЛАТА:  . . . . . . . . . . . . . . . . . . . .     0.0000  141.1760  141.1760";

            //Act
            var result = ParserDetalizationExtensions.ParseNameOfServiceOfBill(text,':');

            //Assert
            Assert.AreEqual("ВАРТІСТЬ ПАКЕТА/ЩОМІСЯЧНА ПЛАТА", result);
        }

        //Test Parsing Contract's Header 
        [Test]
        public void TestParsingHeaderContracts()
        {
            //Arrange                
            List<string> inputed = new List<string>
            {
                "Контракт № 395409092966  Моб.номер: 380500251894",
                "Ціновий Пакет: RED Business M"
            };

            string[] parsers = new string[] {
                @"Владелец",
                @"Контракт №",
                @"Моб.номер",
                @"Ціновий Пакет"
            };


            //Act
            var result = ParserDetalizationExtensions.ParseHeaderOfContractOfBill(inputed, parsers);
            //HeaderOfContractOfBill header = new HeaderOfContractOfBill(contractId, mobileNumber, tarifPackage);

            //Assert
            Assert.AreEqual("395409092966", result.ContractId);
            Assert.AreEqual("+380500251894", result.MobileNumber);
            Assert.AreEqual("RED Business M", result.TarifPackage);
        }

        [Test]
        public void TestGetContractId()
        {
            //Arrange
            string text = "Контракт № 395409092966  Моб.номер: 380500251894";

            //Act
            var result = ParserDetalizationExtensions.GetContractId(text);

            //Assert
            Assert.AreEqual("395409092966", result);
        }

        [Test]
        public void TestGetMobileNumber()
        {
            //Arrange
            string text = "Контракт № 395409092966  Моб.номер: 380500251894";

            //Act
            var result = ParserDetalizationExtensions.GetMobileNumber(text);

            //Assert
            Assert.AreEqual("+380500251894", result);
        }

        [Test]
        public void TestGetTarifPackage()
        {
            //Arrange
            string text = "Ціновий Пакет: RED Business M";

            //Act
            var result = ParserDetalizationExtensions.GetTarifPackage(text);

            //Assert
            Assert.AreEqual("RED Business M", result);
        }



        //Test Convert Internet Trafic
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