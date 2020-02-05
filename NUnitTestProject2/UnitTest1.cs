
namespace NUnitTestProject
{
    using MobileNumbersDetailizationReportGenerator;
    using NUnit.Framework;
    using System;
    using System.Collections.Generic;

    [TestFixture]
    public class Tests
    {

        List<string> listStringsContract;
        List<string> listStringsDetalizationContract;

        [SetUp]
        public void Setup()
        {
            listStringsContract = new List<string>()
            {
                @"�������� � 395383700054  ����� ��������: 380503003378",
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
@"����� ������                        +380505037527 02.06.2019   20:17:47      0:30                0.0000",
@"����� ������.                                5714 02.06.2019   20:18:41         0                0.0000",
@"�����           content ������� SD 5714#Web ������ 02.06.2019   20:18:43         0                2.7451",
@"����� ������                        +380673333305 03.06.2019   09:08:04      3:51                0.0000",
@"�������. ������                       380695222225 27.06.2019   12:12:32      3:41                0.0000",
@"�������. ������ ������� 0800         +380800777950 27.06.2019   12:12:32      3:41                0.0000",
@"������ ������  Vodafone ������     +380952254615 27.06.2019   12:36:57      0:17                0.0000",
@"������ ������  lifecell             +380632888148 27.06.2019   13:37:59      0:17                0.4706",
@"������ ������  Vodafone ������     +380959997593 27.06.2019   15:06:27      2:51                0.0000",
@"����� ������.                                5714 28.06.2019   20:49:01         0                0.0000",
@"�����           content ������� SD 5714#Web ������ 28.06.2019   20:49:04         0                2.7451",
@"����� ������                        +380507770810 29.06.2019   13:11:46      0:24                0.0000",
@"����� ������                        +380503333897 29.06.2019   16:46:10      4:43                0.0000",
@"����� ������.                                5714 29.06.2019   20:52:45         0                0.0000",
@"�����           content ������� SD 5714#Web ������ 29.06.2019   20:52:49         0                2.7451",
@"������ ������  �������             +380688888002 30.06.2019   11:26:10      2:27                1.4118",
@"������ ������  Vodafone ������     +380505577577 30.06.2019   13:51:08      1:34                0.0000",
@"����� ������.                                5714 30.06.2019   20:56:22         0                0.0000",
@"�����           content ������� SD 5714#Web ������ 30.06.2019   20:56:24         0                2.7451",
@"GPRS/CDMA �'��.  �������� �����            internet 01.06.2019   00:00:00  96.00 Mb                0.0000",
@"GPRS/CDMA �'��.  �������� �����            internet 01.06.2019   10:37:07     33 Kb                0.0000",
@"GPRS/CDMA �'��.  �������� �����            internet 02.06.2019   19:38:22    118 Kb                0.0000",
@"GPRS/CDMA �'��.  �������� �����            internet 04.06.2019   19:40:54      1 Kb                0.0000",
@"GPRS/CDMA �'��.  �������� �����            internet 04.06.2019   20:06:40  33.05 Mb                0.0000"
            };
        }


        [Test]
        public void TestParseServicesOfBill()
        {
            var result = ParserDetalizationExtensions.ParseServicesOfBill(listStringsContract);

            Assert.AreEqual(@"���Ҳ��� ������/��̲����� �����", result.Output[0].Name);
            Assert.AreEqual(141.176, result.Output[0].Amount);

            Assert.AreEqual(@"����Ͳ �������-�������", result.Output[5].Name);
            Assert.AreEqual(79.6079, result.Output[5].Amount);
        }

        [Test]
        public void TestParseDetalizationOfContractOfBill()
        {
            var result = ParserDetalizationExtensions.ParseDetalizationOfContractOfBill(listStringsDetalizationContract);

            Assert.Multiple(() =>
            {
                Assert.AreEqual(@"����� ������", result.Output[0].ServiceName);
                Assert.AreEqual(@"02.06.2019", result.Output[0].Date);
                Assert.AreEqual(@"20:17:47", result.Output[0].Time);
                Assert.AreEqual(@"0:30", result.Output[0].DurationA);
                Assert.AreEqual(@"", result.Output[0].DurationB);
                Assert.AreEqual(@"+380505037527", result.Output[0].NumberTarget);
                Assert.AreEqual(@"0.0000", result.Output[0].Cost);

                Assert.AreEqual(@"GPRS/CDMA �'��.  �������� �����", result.Output[21].ServiceName);
                Assert.AreEqual(@"02.06.2019", result.Output[21].Date);
                Assert.AreEqual(@"19:38:22", result.Output[21].Time);
                Assert.AreEqual(@"118 Kb", result.Output[21].DurationA);
                Assert.AreEqual(@"", result.Output[21].DurationB);
                Assert.AreEqual(@"internet", result.Output[21].NumberTarget);
                Assert.AreEqual(@"0.0000", result.Output[21].Cost);
            });
        }

        [Test]
        public void TestParseDetalizationOfContractOfBill_InputNull()
        {
            DetalizationOfContractOfBill excpected = new DetalizationOfContractOfBill();
            var result = ParserDetalizationExtensions.ParseDetalizationOfContractOfBill(null);

            Assert.AreEqual(excpected.ToString(), result.ToString());
        }

        [Test]
        public void TestParseDetalizationOfContractOfBill_InputEmpty()
        {
            DetalizationOfContractOfBill excpected = new DetalizationOfContractOfBill();

            var result = ParserDetalizationExtensions.ParseDetalizationOfContractOfBill(new List<string>());

            Assert.AreEqual(excpected.ToString(), result.ToString());
        }

        [Test]
        public void TestParseStringOfDetalizationOfContractOfBill_InputCorrect()
        {
            string text = @"GPRS/CDMA �'��.  �������� �����            internet 02.06.2019   19:38:22    118 Kb                0.0000";

            var result = ParserDetalizationExtensions.ParseStringOfDetalizationOfContractOfBill(text);

            Assert.Multiple(() => {
                Assert.AreEqual(@"GPRS/CDMA �'��.  �������� �����", result.ServiceName);
                Assert.AreEqual(@"02.06.2019", result.Date);
                Assert.AreEqual(@"19:38:22", result.Time);
                Assert.AreEqual(@"118 Kb", result.DurationA);
                Assert.AreEqual(@"", result.DurationB);
                Assert.AreEqual(@"internet", result.NumberTarget);
                Assert.AreEqual(@"0.0000", result.Cost);
            });
        }

        [Test]
        public void TestParseStringOfDetalizationOfContractOfBill_InputWrong()
        {
            var result = ParserDetalizationExtensions.ParseStringOfDetalizationOfContractOfBill("Wrong or bad Data");

            Assert.AreEqual(null, result);
        }

        [Test]
        public void TestParseStringOfDetalizationOfContractOfBill_InputNull()
        {
            var result = ParserDetalizationExtensions.ParseStringOfDetalizationOfContractOfBill(null);

            Assert.AreEqual(null, result);
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

        [Test]
        public void TestParseNameOfServiceOfBill_MissingParserInText()
        {
            //Arrange
            string text = @"���Ҳ��� ������/��̲����� �����  . . . . . . . . . . . . . . . . . . . .     0.0000  141.1760  141.1760";

            //Act
            var result = ParserDetalizationExtensions.ParseNameOfServiceOfBill(text, ':');

            //Assert
            Assert.AreEqual(@"���Ҳ��� ������/��̲����� �����", result);
        }

        [Test]
        public void TestParseNameOfServiceOfBill_InputNull()
        {
            //Arrange
            string text = null;

            //Act
            var result = ParserDetalizationExtensions.ParseNameOfServiceOfBill(text, ':');

            //Assert
            Assert.AreEqual(null, result);
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
            var result = MultiplierInternetTrafic.ToInternetTrafic(input, multiplier);

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
            var result = MultiplierInternetTrafic.ToInternetTrafic(input, multiplier);

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
            Exception ex = Assert.Throws<Exception>(() => MultiplierInternetTrafic.ToInternetTrafic(input, multiplier));

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