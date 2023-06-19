using CrabExcelDataApp.Util;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Crab_Excel_Data_App_Test.Util
{
    [TestClass]
    public class StringUtilTest
    {
        [TestMethod]
        public void TestNullString()
        {
            var nullStr = StringUtil.ToString(null);

            Assert.IsNotNull(nullStr);
            Assert.AreEqual(string.Empty, nullStr);
        }

        [TestMethod]
        public void TestNornalString()
        {
            var normalStr = StringUtil.ToString("abc");

            Assert.IsNotNull(normalStr);
            Assert.AreEqual("abc", normalStr);
        }
    }
}