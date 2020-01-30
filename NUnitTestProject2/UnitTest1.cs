using MobileNumbersDetailizationReportGenerator;
using NUnit.Framework;
using System;
using System.Collections.Generic;

namespace NUnitTestProject
{
    //   [TestFixture]
    public class Tests
    {

        List<string> listStringsContract;
        List<string> listStringsDetalizationContract;

        [SetUp]
        public void Setup()
        {
            listStringsContract = new List<string>()
            {
                @"�������� � 395381736554  ����� ��������: 380503003348",
                @"�������� �����: RED Business M",
                @"���Ҳ��� ������/��̲����� �����:  . . . . . . . . . . . . . . . . . . . .     0.0000  141.1760  141.1760",
                @"�������, ����Ͳ �� ������ ������: . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .    78.4327",
                @"������ ������ �� �����  . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .    70.5896",
                @"������ ������ � ������ �� ������ . . . . . . . . . . . . . . . . . . . . . . . . . . . . .     7.8431",
                @"�������-�������:  . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .    79.6079",
                @"����Ͳ �������-�������: . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .    79.6079",
                @"SMS\USSD\MMS\�������\����������\������� �� ���� ������ �� ����. ������. . . . . . . . . .    79.6079",
                @"������: . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .   -65.8826",
                @"������ �� ���� ��������� ������� . . . . .. . . . . . . . . . . . . . . . . . . . . . . . . .   -65.8826",
                @"������� �� ���������� (��� ��� �� ��):  . . . . . . . . . . . . . . . . . . . . . . . . . . .   233.3340"
            };
            listStringsDetalizationContract = new List<string>()
            {

            };
        }


        [Test]
        public void TestParseServicesOfBill()
        {
            var result = ParserDetalizationExtensions.ParseServicesOfBill(listStringsContract);

            Assert.AreEqual(@"���Ҳ��� ������/��̲����� �����", result.Output[0].Name);
            Assert.AreEqual(141.176, result.Output[0].Amount);

            Assert.AreEqual(@"������� �� ���������� (��� ��� �� ��)", result.Output[5].Name);
            Assert.AreEqual(233.334, result.Output[5].Amount);
        }

        [Test]
        public void TestParseDetalizationOfContractOfBill()
        {
            var result = ParserDetalizationExtensions.ParseDetalizationOfContractOfBill(listStringsDetalizationContract);

            Assert.AreEqual(@"���Ҳ��� ������/��̲����� �����", result.Output[0].Name);
            Assert.AreEqual(141.176, result.Output[0].Amount);

            Assert.AreEqual(@"������� �� ���������� (��� ��� �� ��)", result.Output[5].Name);
            Assert.AreEqual(233.334, result.Output[5].Amount);

        }

        [Test]
        public void TestParseCostOfServiceOfBill()
        {
            //Arrange
            string text = @"���Ҳ��� ������/��̲����� �����:  . . . . . . . . . . . . . . . . . . . .     0.0000  141.1760  141.1760";

            //Act
            var result = ParserDetalizationExtensions.ParseCostOfServiceOfBill(text);

            //Assert
            Assert.AreEqual(141.1760, result);
        }

        [Test]
        public void TestParseCostOfServiceOfBill_WithWrongInput()
        {
            //Arrange
            string text = @"���Ҳ��� ������/��̲����� �����:  . . . . . . . . . . . . . . . . . . . .     ";

            //Act
            var result = ParserDetalizationExtensions.ParseCostOfServiceOfBill(text);

            //Assert
            Assert.AreEqual(0, result);
        }

        [Test]
        public void TestParseNameOfServiceOfBill()
        {
            //Arrange
            string text = @"���Ҳ��� ������/��̲����� �����:  . . . . . . . . . . . . . . . . . . . .     0.0000  141.1760  141.1760";

            //Act
            var result = ParserDetalizationExtensions.ParseNameOfServiceOfBill(text, ':');

            //Assert
            Assert.AreEqual("���Ҳ��� ������/��̲����� �����", result);
        }

        //Test Parsing Contract's Header 
        [Test]
        public void TestParsingHeaderContracts()
        {
            //Arrange                
            List<string> inputed = new List<string>
            {
                "�������� � 395409092966  ���.�����: 380500251894",
                "ֳ����� �����: RED Business M"
            };

            string[] parsers = new string[] {
                @"��������",
                @"�������� �",
                @"���.�����",
                @"ֳ����� �����"
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
            string text = "�������� � 395409092966  ���.�����: 380500251894";

            //Act
            var result = ParserDetalizationExtensions.GetContractId(text);

            //Assert
            Assert.AreEqual("395409092966", result);
        }

        [Test]
        public void TestGetMobileNumber()
        {
            //Arrange
            string text = "�������� � 395409092966  ���.�����: 380500251894";

            //Act
            var result = ParserDetalizationExtensions.GetMobileNumber(text);

            //Assert
            Assert.AreEqual("+380500251894", result);
        }

        [Test]
        public void TestGetTarifPackage()
        {
            //Arrange
            string text = "ֳ����� �����: RED Business M";

            //Act
            var result = ParserDetalizationExtensions.GetTarifPackage(text);

            //Assert
            Assert.AreEqual("RED Business M", result);
        }



        //Test Convertor Internet Trafic
        //Arrange
        //correct Data
        [TestCase("719 Gb", "Gb", 719)]
        [TestCase("719 Gb", "GB", 719)]
        [TestCase("200 Mb", "Mb", 200)]
        [TestCase("200 Mb", "MB", 200)]
        [TestCase("300 Kb", "Kb", 300)]
        [TestCase("300 Kb", "KB", 300)]
        [TestCase("120 b", "b", 120)]
        [TestCase("120 b", "B", 120)]
        [TestCase("201Mb", "b", 210763776)]
        [TestCase("250 Mb", "Kb", 256000)]
        [TestCase("300 Kb", "MB", 0.293)]
        public void TestToInternetTrafic(string input, string multiplier, double output)
        {
            //Act
            var result = WinFormsExtensions.ToInternetTrafic(input, multiplier);

            //Assert
            Assert.AreEqual(output, result);
        }

        //wrong inputed data
        //Arrange
        [TestCase("garbage", "Mb", 0)]
        [TestCase("", "Mb", 0)]
        [TestCase("kolomb", "Mb", 0)]
        public void TestToInternetTrafic_With_WrongData(string input, string multiplier, double output)
        {
            //Act
            var result = WinFormsExtensions.ToInternetTrafic(input, multiplier);

            //Assert
            Assert.AreEqual(output, result);
        }

        //wrong multiplier
        //Arrange
        [TestCase("kolomb", "rd")]
        [TestCase("200 Mb", "md")]
        public void TestToInternetTrafic_With_WrongMultiplier(string input, string multiplier)
        {
            //Act
            Exception ex = Assert.Throws<Exception>(() => WinFormsExtensions.ToInternetTrafic(input, multiplier));

            //Assert
            Assert.AreEqual("Wrong multiplier!", ex.Message);

            //2nd way to check
            // Assert.That(() =>
            // {
            //    WinFormsExtensions.ToInternetTrafic(input, multiplier);
            //}, Throws.TypeOf<Exception>().With.Message.EqualTo("Wrong multiplier!"));
        }

    }
}