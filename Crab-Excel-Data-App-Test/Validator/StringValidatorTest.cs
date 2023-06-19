using CrabExcelDataApp.Validator;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Crab_Excel_Data_App_Test.Validator
{
    [TestClass]
    public class StringValidatorTest
    {
        [TestMethod]
        public void TestIsBlank()
        {
            Assert.IsFalse(StringValidator.IsBlank("a"));
            Assert.IsFalse(StringValidator.IsBlank("d "));
            Assert.IsFalse(StringValidator.IsBlank("c  "));
            Assert.IsFalse(StringValidator.IsBlank("  b"));
            Assert.IsFalse(StringValidator.IsBlank("  e "));

            Assert.IsTrue(StringValidator.IsBlank(null));
            Assert.IsTrue(StringValidator.IsBlank(""));
            Assert.IsTrue(StringValidator.IsBlank(" "));
            Assert.IsTrue(StringValidator.IsBlank("                     "));
        }
    }
}
