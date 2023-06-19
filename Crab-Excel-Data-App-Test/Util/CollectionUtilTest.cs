using CrabExcelDataApp.Util;
using global::System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace Crab_Excel_Data_App_Test.Util
{
    [TestClass]
    public class CollectionUtilTest
    {
        [TestMethod]
        public void TestNotNullList()
        {
            var list = new List<string>();
            list.Add("1");
            list.Add("2");
            Assert.IsFalse(CollectionUtil.IsNullOrEmpty(list));

            var list2 = new string[] { "a", "b" };
            Assert.IsFalse(CollectionUtil.IsNullOrEmpty(list2));

            var list3 = new Dictionary<string, string>();
            list3.Add("key", "1");
            list3.Add("name", "Ben");
            Assert.IsFalse(CollectionUtil.IsNullOrEmpty(list3));
        }

        [TestMethod]
        public void TestNullList()
        {
            var list = new List<string>();
            Assert.IsTrue(CollectionUtil.IsNullOrEmpty(list));

            var list2 = new string[] { };
            Assert.IsTrue(CollectionUtil.IsNullOrEmpty(list2));

            var list3 = new Dictionary<string, string>();
            Assert.IsTrue(CollectionUtil.IsNullOrEmpty(list3));

            List<string> list4 = null;
            Assert.IsTrue(CollectionUtil.IsNullOrEmpty(list4));
        }
    }
}
