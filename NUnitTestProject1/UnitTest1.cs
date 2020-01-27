using NUnit.Framework;
using MobileNumbersDetailizationReportGenerator;

namespace NUnitTestProject1
{
    public class Tests
    {
        [Test]
        public void TestToInternetTrafic_200_Mb_Wait_200()
        {
            string text = "200 Mb";
            var result = WinFormsExtensions.ToInternetTrafic(text, "Mb");

           Assert.Equals(200, result);
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